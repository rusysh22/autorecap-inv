import os
import pandas as pd
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from datetime import datetime
import io
import base64
import json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

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
            "任务单号 kode tugas", "id"
        ]
        
        aliases_nama = [
            "rute", "ritase", "kode ritase", "nama rute", "nama tugas", 
            "线路 rute"
        ]
        
        # 3. Find Matching Columns
        kode_col_name = None
        for alias in aliases_kode:
            if alias in clean_headers:
                kode_col_name = clean_headers[alias]
                break
                
        nama_col_name = None
        for alias in aliases_nama:
            if alias in clean_headers:
                nama_col_name = clean_headers[alias]
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
            
            # Apply cleaning to 'Nama Tugas' from col G(6)
            # LOGIC: If master_mapping exists and Kode Tugas matches, use Master Value.
            # ELSE: Use clean_route_name logic.
            
            def resolve_route_name(row):
                original_code = str(row.iloc[3]).strip() # Col D (Kode Tugas)
                original_name = row.iloc[6] # Col G (Nama Tugas)
                
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

            temp_df['Nama Tugas'] = df.apply(resolve_route_name, axis=1)
            
            # Map columns actually found in source
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
            temp_df['Total pembayaran aktual'] = df.iloc[:, 22] # Col W
            
            # Filter out empty rows (e.g. if 'Agen Operasional' or 'Kode Tugas' is empty)
            temp_df = temp_df.dropna(subset=['Agen Operasional', 'Kode Tugas'])
            
            # Tag with source filename (for frontend display only)
            temp_df['source_file'] = filename_display

            # Calculate summary for this file
            file_total = temp_df['Total pembayaran aktual'].sum()
            ppn_total = temp_df['PPN'].sum()
            pph_total = temp_df['PPH'].sum()
            
            file_summaries.append({
                "filename": filename_display,
                "rows": len(temp_df),
                "amount": float(file_total),
                "ppn": float(ppn_total),
                "pph": float(pph_total),
                "status": "Success"
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
    if missing_lookup_codes:
        count = len(missing_lookup_codes)
        sample = list(missing_lookup_codes)[:3]
        msg = f"⚠️ {count} Kode Tugas not found in Master Data (using default name). Example: {', '.join(sample)}"
        if count > 3: msg += "..."
        warnings.add(msg)
        
    return final_df, file_summaries, list(warnings)

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
    final_df, file_summaries, warnings = process_excel_files(files, master_mapping=master_mapping)
    
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
        "warnings": all_warnings
    })

if __name__ == '__main__':
    app.run(debug=True)
