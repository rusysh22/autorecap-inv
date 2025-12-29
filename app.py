import os
import pandas as pd
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from datetime import datetime
import io
import base64

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

def process_excel_files(files):
    """
    Processes a list of file storages objects (in-memory).
    Returns:
        final_df (DataFrame): Consolidated data
        file_summaries (list): Statistics per file
    """
    all_data = []
    file_summaries = []

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
        
    return final_df, file_summaries

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

    # Process Files (In-Memory)
    final_df, file_summaries = process_excel_files(files)
    
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
            worksheet.row_dimensions[1].height = 30
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
        "display_columns": excel_columns
    })

if __name__ == '__main__':
    app.run(debug=True)
