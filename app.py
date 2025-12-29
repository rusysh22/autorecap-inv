from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import glob
from datetime import datetime

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- HELPER FUNCTIONS ---

def clean_route_name(route_str):
    """
    Cleans the 'Nama Tugas' from source text.
    Logic: Split by '-', take first 3 chars of each segment, join with ' - '.
    Example: "PTI777-BGR999..." -> "PTI - BGR"
    """
    if not isinstance(route_str, str):
        return ""
    
    parts = route_str.split('-')
    cleaned_parts = []
    
    for part in parts:
        cleaned = part.strip()[:3]
        if cleaned:
            cleaned_parts.append(cleaned)
            
    return " - ".join(cleaned_parts)

def process_excel_files(file_paths):
    all_data = []
    file_summaries = []
    
    for file_path in file_paths:
        try:
            filename_display = os.path.basename(file_path).split('_', 1)[1] if '_' in os.path.basename(file_path) else os.path.basename(file_path)
            
            # Read Excel, Header at Row 4 (index 3)
            df = pd.read_excel(file_path, header=3)
            
            # Verify structure (check if we have enough columns, max index needed is 22)
            if df.shape[1] < 23:
                print(f"Skipping {file_path}: Not enough columns (found {df.shape[1]})")
                file_summaries.append({
                    "filename": filename_display,
                    "rows": 0,
                    "amount": 0,
                    "status": "Error: Columns"
                })
                continue

            # Create a localized dataframe for this file
            temp_df = pd.DataFrame()
            
            # Mapping based on Screenshot (v2)
            temp_df['Agen Operasional'] = df.iloc[:, 1] # Col B
            temp_df['Kode Tugas'] = df.iloc[:, 3]       # Col D
            
            # Apply cleaning to 'Nama Tugas' from col G(6)
            temp_df['Nama Tugas'] = df.iloc[:, 6].apply(clean_route_name)
            
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
            print(f"Error processing {file_path}: {e}")
            file_summaries.append({
                "filename": os.path.basename(file_path),
                "rows": 0,
                "amount": 0,
                "status": "Error"
            })

    final_df = pd.DataFrame()
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
    return final_df, file_summaries

# --- ROUTES ---

@app.route('/')
def index():
    return render_template('index.html')

import time

# ... (existing imports)

# --- HELPER FUNCTIONS ---

def cleanup_folders():
    """
    Removes files from UPLOAD_FOLDER and OUTPUT_FOLDER that are older than 1 hour.
    Prevents accumulation of 'junk' files.
    """
    now = time.time()
    cutoff = now - 3600 # 1 hour ago
    
    for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
        if not os.path.exists(folder):
            continue
            
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path):
                    file_mtime = os.path.getmtime(file_path)
                    if file_mtime < cutoff:
                        os.remove(file_path)
                        print(f"Deleted old file: {filename}")
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

# ... (existing clean_route_name and process_excel_files functions)

@app.route('/api/process', methods=['POST'])
def process_files():
    # 1. Run Cleanup First
    cleanup_folders()

    # Clear upload folder first (optional, or manage better)
    uploaded_files = request.files.getlist('files')
    # ...
    filename_suffix = request.form.get('filename_suffix', '').strip()
    
    saved_paths = []
    
    if not uploaded_files:
        return jsonify({"error": "No files uploaded"}), 400

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    for file in uploaded_files:
        if file.filename:
            filename = f"{timestamp}_{file.filename}"
            path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(path)
            saved_paths.append(path)
    
    # Process
    final_df, file_summaries = process_excel_files(saved_paths)
    
    # Define exact column order for Excel
    excel_columns = [
        'Agen Operasional', 'Kode Tugas', 'Nama Tugas', 'Plat Mobil', 
        'Jenis Kendaraan', 'Mode Operasi', 'Metode Perhitungan', 'Berat', 
        'Tarif Pengiriman per kg', 'Tarif Pengiriman Sistem', 'PPN', 'PPH', 
        'Total pembayaran aktual'
    ]
    
    # Reorder if not empty, otherwise create empty with these columns
    if not final_df.empty:
        # Ensure all Excel columns exist
        for col in excel_columns:
            if col not in final_df.columns:
                final_df[col] = "" 
        
        # Keep source_file if it exists, otherwise strict filter
        cols_to_keep = excel_columns + (['source_file'] if 'source_file' in final_df.columns else [])
        final_df = final_df[cols_to_keep]
    else:
        final_df = pd.DataFrame(columns=excel_columns)
    
    # Construct Output Filename with Prefix
    # Prefix: 陆运数据核对
    if filename_suffix:
        # User provided suffix
        output_filename = f"陆运数据核对 {filename_suffix}.xlsx"
    else:
        # Default fallback if no suffix provided
        output_filename = f"陆运数据核对 {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    
    try:
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write ONLY the Excel columns, excluding internal helpers like 'source_file'
            final_df[excel_columns].to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Access the workbook and sheet
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Define Styles
            header_font = Font(name='SimSun', size=11, color="FF0000", bold=True)
            header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            thin_border = Border(left=Side(style='thin'), 
                                 right=Side(style='thin'), 
                                 top=Side(style='thin'), 
                                 bottom=Side(style='thin'))
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            # Apply to Header Row
            worksheet.row_dimensions[1].height = 30 # Set row height
            
            # Special Font for Black Headers
            black_header_font = Font(name='SimSun', size=11, color="000000", bold=True)

            for cell in worksheet[1]:
                if cell.value in ['Berat', 'Tarif Pengiriman per kg']:
                     cell.font = black_header_font
                else:
                     cell.font = header_font
                
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = center_alignment
                
            # Auto-adjust column widths
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

    except Exception as e:
        print(f"Error styling excel: {e}")
        final_df.to_excel(output_path, index=False)
    
    # Convert to JSON for preview (limit 100 rows)
    preview_data = final_df.head(100).to_dict(orient='records')
    
    # Calculate Summary
    summary = {
        "total_files": len(saved_paths),
        "total_rows": len(final_df),
        "total_amount": float(final_df['Total pembayaran aktual'].sum()) if not final_df.empty else 0,
        "download_url": f"/api/download?filename={output_filename}",
        "output_filename": output_filename,
        "file_details": file_summaries
    }

    return jsonify({
        "success": True,
        "data": preview_data,
        "display_columns": excel_columns,
        "summary": summary
    })

@app.route('/api/download')
def download():
    filename = request.args.get('filename')
    path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "File not found", 404

if __name__ == '__main__':
    app.run(debug=True)
