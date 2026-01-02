import os
import pandas as pd
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from datetime import datetime
import io
import base64
import json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB limit (approx 15 files * 5MB + overhead)

# --- HELPER FUNCTIONS ---

def clean_route_name(route_name):
    """
    Cleans 'Nama Tugas' (Route Name).
    Logic: Split by '-', take first 3 parts, join with '-'.
    Example: 'BGR-SOC-A001-X' -> 'BGR-SOC-A001'
    """
    if not isinstance(route_name, str):
        return str(route_name)
    parts = route_name.split('-')
    if len(parts) >= 3:
        return "-".join(parts[:3])
    return route_name

def load_master_data(file_storage):
    """
    Parses the Master Data Excel file.
    Expected columns: 'Kode Tugas', 'Nama Tugas' (fuzzy match based on aliases).
    Returns: Dict mapping {Kode Tugas (str): Nama Tugas (str)}
    """
    try:
        df = pd.read_excel(file_storage, engine='openpyxl')
        
        # 1. Cleaning Headers: Convert to lower, strip whitespace, replace newlines with space
        # We keep a map of {clean_header: original_header} to reference data later
        clean_headers = {
            str(col).lower().replace('\n', ' ').strip(): col 
            for col in df.columns
        }
        
        # 2. Define Aliases (Normalized to lower case)
        aliases_kode = [
            "kode", "tugas id", "kode tugas", 
            "任务单号", "id"
        ]
        
        aliases_nama = [
            "rute", "ritase", "kode ritase", "nama rute", "nama tugas", 
            "线路"
        ]
        
        # 3. Find Matching Columns
        kode_col_name = None
        for clean_header in clean_headers:
            if any(alias in clean_header for alias in aliases_kode):
                kode_col_name = clean_headers[clean_header]
                break
                
        nama_col_name = None
        for clean_header in clean_headers:
            if any(alias in clean_header for alias in aliases_nama):
                nama_col_name = clean_headers[clean_header]
                break
        
        # 4. Validation
        if not kode_col_name or not nama_col_name:
            # Return specific error indicator
            return {"__error__": f"Master Data Error: Columns not found in {getattr(file_storage, 'filename', 'file')}. Found: {list(clean_headers.keys())}"}
            
        # 5. Create mapping
        # Convert Kode Tugas to string and strip for reliable matching
        mapping = pd.Series(
            df[nama_col_name].values, 
            index=df[kode_col_name].astype(str).str.strip()
        ).to_dict()
        
        return mapping
    except Exception as e:
        return {"__error__": f"Error loading master data: {str(e)}"}

def process_excel_files(files, master_mapping=None):
    """
    Processes a list of file storages objects (in-memory).
    Returns:
        final_df (DataFrame): Consolidated data
        file_summaries (list): Statistics per file
        warnings (list): List of warning messages
    """
    all_data = []
    file_summaries = []
    warnings = set() # Use set to avoid duplicate warnings
    
    missing_lookup_codes = set() # Track codes not found in Master Data

    for file in files:
        try:
            filename_display = secure_filename(file.filename)
            
            # Read directly from memory
            # engine='openpyxl' works with file-like objects
            df = pd.read_excel(file, engine='openpyxl', header=3)
            
            # Basic validation: Check if required columns exist by index
            # We strictly need up to index 22 (Col W)
            if df.shape[1] < 23:
                file_summaries.append({
                    "filename": filename_display,
                    "rows": 0,
                    "amount": 0,
                    "ppn": 0,
                    "pph": 0,
                    "status": "Error: Columns"
                })
                continue

            # Create a localized dataframe for this file
            temp_df = pd.DataFrame()
            
            # Mapping based on Screenshot (v2)
            temp_df['Agen Operasional'] = df.iloc[:, 1] # Col B
            temp_df['Kode Tugas'] = df.iloc[:, 3]       # Col D
            temp_df['Total pembayaran aktual'] = df.iloc[:, 22] # Col W
            temp_df['Raw_Nama_Tugas'] = df.iloc[:, 6]   # Col G (Needed for resolution)
            
            # --- ROW CLEANING STEP 1 ---
            # Remove rows where crucial keys are missing immediately
            # Convert to string, strip, and coerce empty ('nan', 'none', '') to NaN
            temp_df['Agen Operasional'] = temp_df['Agen Operasional'].astype(str).str.strip().replace(['nan', 'NaN', 'None', '', 'NaT'], float('nan'))
            temp_df['Kode Tugas'] = temp_df['Kode Tugas'].astype(str).str.strip().replace(['nan', 'NaN', 'None', '', 'NaT'], float('nan'))
            
            # Drop purely empty rows
            temp_df = temp_df.dropna(subset=['Agen Operasional', 'Kode Tugas'])
            
            # --- ROW CLEANING STEP 2 ---
            # Filter out known Footer/Anomaly keywords from 'Kode Tugas'
            anomaly_keywords = ['dicek oleh', 'diketahui oleh', 'dibuatkan', 'disetujui oleh', 'bill periode', 'total', 'print date']
            
            def is_anomaly(row):
                val = str(row['Kode Tugas']).lower()
                # 1. Check for keywords
                if any(kw in val for kw in anomaly_keywords):
                    return True
                # 2. Check for Colon (common in labels like "Note :", "Oleh :")
                if val.endswith(' :') or val.endswith(':'):
                    return True
                
                # 3. Check Payment Amount (Signature rows usually have empty/NaN Total)
                total_val = row['Total pembayaran aktual']
                is_total_empty = pd.isna(total_val) or str(total_val).strip() == ''
                
                # If Total is empty, it's likely a footer/signature layout row
                if is_total_empty:
                    return True
                    
                return False

            temp_df = temp_df[~temp_df.apply(is_anomaly, axis=1)]

            # Apply cleaning to 'Nama Tugas' 
            # LOGIC: Resolve using Master Data or clean fallback on 'Raw_Nama_Tugas'
            
            def resolve_route_name(row):
                original_code = str(row['Kode Tugas']).strip()
                original_name = row['Raw_Nama_Tugas']
                
                # Skip valid lookup check if Kode Tugas is empty/nan
                if not original_code or original_code.lower() == 'nan':
                     return clean_route_name(original_name)

                if master_mapping:
                    if original_code in master_mapping:
                        return master_mapping[original_code]
                    else:
                        # Track missing lookup
                        missing_lookup_codes.add(original_code)
                
                return clean_route_name(original_name)

            # Apply on filtered temp_df instead of original df
            temp_df['Nama Tugas'] = temp_df.apply(resolve_route_name, axis=1)
            
            # Map remaining columns (Using index from original df via loc if needed? 
            # No, we need to pull them from df.iloc but aligned to temp_df index)
            # Since temp_df is a slice/copy with same index, we can just assign by index matching logic 
            # OR better: Add these columns to temp_df BEFORE filtering if we want to be safe,
            # BUT adding them now works because pandas aligns by index.
            # However, `df.iloc[:, 7]` is a Series with full index. temp_df has subset index. 
            # Direct assignment `temp_df['Col'] = df.iloc[...]` works on index alignment.
            
            temp_df['Plat Mobil'] = df.iloc[:, 7]       # Col H
            temp_df['Jenis Kendaraan'] = df.iloc[:, 8]  # Col I
            temp_df['Mode Operasi'] = df.iloc[:, 9].astype(str).str.lower()     # Col J
            temp_df['Metode Perhitungan'] = df.iloc[:, 14].astype(str).str.lower().str.replace('per/', '') # Col O
            
            # Defaults for missing columns
            temp_df['Berat'] = "" 
            temp_df['Tarif Pengiriman per kg'] = "" 
            
            temp_df['Tarif Pengiriman Sistem'] = df.iloc[:, 15] # Col P
            temp_df['PPN'] = df.iloc[:, 20] # Col U
            temp_df['PPH'] = df.iloc[:, 21] # Col V
            # Total payment is already in temp_df, but ensure we keep it
            
            # Drop the auxiliary columns if not needed in final output
            # (Raw_Nama_Tugas is not needed in final)
            temp_df = temp_df.drop(columns=['Raw_Nama_Tugas'])
            
            # Tag with source filename (for frontend display only)
            temp_df['source_file'] = filename_display


            
            # Helper for safe float conversion
            def safe_float_convert(val):
                try:
                    if pd.isna(val): return 0.0
                    
                    # If it's already a number, return it
                    if isinstance(val, (int, float)):
                        return float(val)
                        
                    s = str(val).strip().replace('Rp', '').replace('IDR', '').strip()
                    
                    # Indonesian Format Handling:
                    # 1.100.000 -> 1100000
                    # 1.100.000,50 -> 1100000.50
                    # 1,5 -> 1.5
                    
                    # Logic: 
                    # If multiple dots exist, they are thousand separators -> remove them
                    # If one dot exists and one comma exists -> decide by position
                    # If only dot exists: potentially thousand separator OR decimal. 
                    #    - Heuristic: if 3 digits after dot, treat as thousand? No, dangerous.
                    #    - Given checking PPH/PPN (currency), usually no decimals or comma decimal.
                    # Let's assume standard ID format: Strip all dots, replace comma with dot.
                    
                    if '.' in s and ',' in s:
                        if s.rfind(',') > s.rfind('.'): # 1.234,56
                            s = s.replace('.', '').replace(',', '.')
                        else: # 1,234.56 (US, unlikely but possible)
                            s = s.replace(',', '')
                    elif '.' in s:
                         # Case: 1.234 (could be 1234 or 1.234)
                         # If index 22 is amount, PPH/PPN usually similar scale
                         # Safest for ID currency: Remove dots (thousand sep)
                         # But what if value is 1.5%? (0.015). 
                         # Excel usually gives float for that. String usually implies formatting.
                         # Try removing dots if it looks like thousand sep (e.g. 14.500)
                         # Simple removal of dots for integer-like strings
                         s = s.replace('.', '')
                    elif ',' in s:
                        s = s.replace(',', '.')
                        
                    return float(s)
                except Exception as e:
                    # print(f"Float conv error for {val}: {e}")
                    return None

            # Iterate through rows for validation
            for idx, row in temp_df.iterrows():
                row_num = idx + 1
                
                # Check 1: Jenis Mobil
                jenis_mobil = str(row.get('Jenis Kendaraan', '')).strip().upper()
                if jenis_mobil and not ('CDDL' in jenis_mobil or 'TWB' in jenis_mobil):
                    file_anomalies.append(f"Row {row_num}: Jenis Mobil is '{jenis_mobil}' (Expected: 'CDDL' or 'TWB')")
                
                # Check 2: PPH (Tax) -> Warning if Positive
                pph_raw = row.get('PPH', 0)
                pph_val = safe_float_convert(pph_raw)
                
                if pph_val is None:
                     pass 
                elif pph_val > 0:
                     # print(f"DEBUG: Anomaly Row {row_num} PPH Positive: {pph_val} (Raw: {pph_raw})")
                     file_anomalies.append(f"Row {row_num}: PPH is {pph_val:,.0f} (Expected: Negative)")
                    
                # Check 3: PPN (VAT) -> Warning if Negative
                ppn_raw = row.get('PPN', 0)
                ppn_val = safe_float_convert(ppn_raw)
                
                if ppn_val is not None and ppn_val < 0:
                    # print(f"DEBUG: Anomaly Row {row_num} PPN Negative: {ppn_val} (Raw: {ppn_raw})")
                    file_anomalies.append(f"Row {row_num}: PPN is {ppn_val:,.0f} (Expected: Positive)")

            # Calculate summary for this file
            try:
                file_total = temp_df['Total pembayaran aktual'].apply(safe_float_convert).sum()
                ppn_total = temp_df['PPN'].apply(safe_float_convert).sum()
                pph_total = temp_df['PPH'].apply(safe_float_convert).sum()
            except:
                file_total = 0
                ppn_total = 0
                pph_total = 0
            
            status_label = "Success"
            if file_anomalies:
                status_label = "Warning"
                print(f"File {filename_display} has {len(file_anomalies)} anomalies.")
            
            file_summaries.append({
                "filename": filename_display,
                "rows": len(temp_df),
                "amount": float(file_total),
                "ppn": float(ppn_total),
                "pph": float(pph_total),
                "status": status_label,
                "anomalies": file_anomalies
            })
            
            all_data.append(temp_df)
            
        except Exception as e:
            print(f"Error processing file: {e}")
            file_summaries.append({
                "filename": getattr(file, 'filename', 'Unknown'),
                "rows": 0,
                "amount": 0,
                "status": "Error"
            })

    final_df = pd.DataFrame()
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
    
    # --- GLOBAL VALIDATIONS ---
    
    # 1. Negative Total
    if not final_df.empty:
        total_amount = final_df['Total pembayaran aktual'].sum()
        if total_amount < 0:
            warnings.add(f"⚠️ Total Amount is Negative: {total_amount:,.0f}. Please check column placement or source data.")

    # 2. Missing Lookups Summary
    # We will pass the raw list to frontend instead of formatting a string here
    missing_codes_list = list(missing_lookup_codes)
        
    return final_df, file_summaries, list(warnings), missing_codes_list

# --- ROUTES ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process', methods=['POST'])
def process_files():
    if 'files' not in request.files:
        return jsonify({"success": False, "error": "No files uploaded"}), 400
    
    files = request.files.getlist('files')
    filename_suffix = request.form.get('filename_suffix', '').strip()
    
    if not files or files[0].filename == '':
        return jsonify({"success": False, "error": "No files selected"}), 400

    # Handle Master Data Sources
    master_mapping = {}
    master_errors = []
    
    # 1. Multiple Files
    if 'master_files' in request.files:
        master_files_list = request.files.getlist('master_files')
        for m_file in master_files_list:
            if m_file and m_file.filename != '':
                print(f"Processing Master File: {m_file.filename}")
                file_mapping = load_master_data(m_file)
                
                # Check for errors in loading
                if "__error__" in file_mapping:
                    master_errors.append(file_mapping["__error__"])
                else:
                    master_mapping.update(file_mapping)
                
    # 2. JSON Data (Pasted)
    master_json_str = request.form.get('master_data_json')
    if master_json_str:
        try:
            pasted_data = json.loads(master_json_str) # List of dicts {kode, nama}
            for item in pasted_data:
                k = str(item.get('kode', '')).strip()
                v = str(item.get('nama', '')).strip()
                if k and v:
                    master_mapping[k] = v
            print(f"Merged {len(pasted_data)} records from paste.")
        except Exception as e:
            print(f"Error parsing master_data_json: {e}")
            master_errors.append("Error parsing pasted Master Data.")
            
    # If we have critical master data errors, we might want to stop or warn
    # For now, let's pass them to frontend
    
    print(f"Total Master Records Loaded: {len(master_mapping)}")

    # Process Files (In-Memory)
    final_df, file_summaries, warnings, missing_codes = process_excel_files(files, master_mapping=master_mapping)
    
    # Prepend master errors to warnings
    all_warnings = master_errors + warnings
    
    # Define Columns
    excel_columns = [
        'Agen Operasional', 'Kode Tugas', 'Nama Tugas', 'Plat Mobil', 
        'Jenis Kendaraan', 'Mode Operasi', 'Metode Perhitungan', 
        'Berat', 'Tarif Pengiriman per kg', 'Tarif Pengiriman Sistem', 
        'PPN', 'PPH', 'Total pembayaran aktual'
    ]
    
    # Filter final_df
    if not final_df.empty:
        cols_to_keep = excel_columns + (['source_file'] if 'source_file' in final_df.columns else [])
        final_df = final_df[cols_to_keep]
    else:
        final_df = pd.DataFrame(columns=excel_columns)
    
    # Construct Output Filename
    if filename_suffix:
        output_filename = f"陆运数据核对 {filename_suffix}.xlsx"
    else:
        output_filename = f"陆运数据核对 {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        
    # Generate Excel in Memory
    output_io = io.BytesIO()
    try:
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        
        with pd.ExcelWriter(output_io, engine='openpyxl') as writer:
            # Write ONLY the Excel columns
            final_df[excel_columns].to_excel(writer, index=False, sheet_name='Sheet1')
            
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Styles
            header_font = Font(name='SimSun', size=11, color="FF0000", bold=True)
            black_header_font = Font(name='SimSun', size=11, color="000000", bold=True)
            header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            # Apply to Header
            worksheet.row_dimensions[1].height = 40
            for cell in worksheet[1]:
                if cell.value in ['Berat', 'Tarif Pengiriman per kg']:
                     cell.font = black_header_font
                else:
                     cell.font = header_font
                     
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = center_alignment
                
            # Auto-adjust widths
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

        # Seek to beginning
        output_io.seek(0)
        
        # Encode to Base64
        excel_base64 = base64.b64encode(output_io.getvalue()).decode('utf-8')

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

    # JSON Response
    data_preview = final_df.replace({float('nan'): None}).to_dict(orient='records')
    
    summary = {
        "total_files": len(files),
        "total_rows": len(final_df),
        "total_amount": float(final_df['Total pembayaran aktual'].sum()) if not final_df.empty else 0,
        "output_filename": output_filename,
        "excel_data": excel_base64, # Base64 encoded file
        "file_details": file_summaries
    }
    
    return jsonify({
        "success": True, 
        "data": data_preview, 
        "summary": summary,
        "display_columns": excel_columns,
        "warnings": all_warnings,
        "missing_codes": missing_codes
    })

if __name__ == '__main__':
    app.run(debug=True)
