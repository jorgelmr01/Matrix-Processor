#!/usr/bin/env python3
"""
Matrix Processor - A web app for creating intersection matrices from Excel/CSV files.
Run with: python app.py
"""

import os
import sys
import json
import webbrowser
import threading
import traceback
from datetime import datetime
from io import BytesIO
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse

# ============================================
# LOGGING SYSTEM
# ============================================
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'matrix_processor_debug.log')

def log_debug(message, data=None):
    """Write detailed debug information to log file"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"\n[{timestamp}] {message}\n")
        if data is not None:
            if isinstance(data, (dict, list)):
                f.write(json.dumps(data, indent=2, default=str)[:5000] + "\n")
            else:
                f.write(str(data)[:5000] + "\n")

def log_error(message, exception=None):
    """Write error information to log file"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"\n[{timestamp}] ERROR: {message}\n")
        if exception:
            f.write(f"Exception: {str(exception)}\n")
            f.write(f"Traceback:\n{traceback.format_exc()}\n")

def clear_log():
    """Clear the log file"""
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        f.write(f"=== Matrix Processor Debug Log ===\n")
        f.write(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Log file: {LOG_FILE}\n")
        f.write("=" * 50 + "\n")

# Check for required packages and install if missing
def check_dependencies():
    required = ['openpyxl', 'pandas']
    missing = []
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    
    if missing:
        print(f"Installing required packages: {', '.join(missing)}")
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing, 
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("Packages installed successfully!")

check_dependencies()

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# Store uploaded files and processing state
app_state = {
    'files': {},
    'file_data': [],
    'filter_file': None
}

class MatrixProcessorHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=os.path.dirname(os.path.abspath(__file__)), **kwargs)
    
    def log_message(self, format, *args):
        pass  # Suppress logging
    
    def send_json(self, data, status=200):
        response = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(response))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(response)
    
    def send_file_download(self, data, filename):
        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.send_header('Content-Length', len(data))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(data)
    
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    
    def do_GET(self):
        if self.path == '/':
            self.path = '/index.html'
        
        if self.path == '/api/status':
            self.send_json({'status': 'ok', 'files': len(app_state['file_data'])})
        elif self.path == '/api/debug-log':
            # Return the debug log file contents
            try:
                with open(LOG_FILE, 'r', encoding='utf-8') as f:
                    log_content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'text/plain; charset=utf-8')
                self.send_header('Content-Length', len(log_content.encode('utf-8')))
                self.end_headers()
                self.wfile.write(log_content.encode('utf-8'))
            except Exception as e:
                self.send_json({'error': str(e)}, 500)
        else:
            super().do_GET()
    
    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        
        if self.path == '/api/upload':
            self.handle_upload(content_length)
        elif self.path == '/api/upload-filter':
            self.handle_upload_filter(content_length)
        elif self.path == '/api/process':
            self.handle_process(content_length)
        elif self.path == '/api/compute':
            self.handle_compute(content_length)
        elif self.path == '/api/export':
            self.handle_export(content_length)
        elif self.path == '/api/validate-filter':
            self.handle_validate_filter(content_length)
        elif self.path == '/api/reset':
            app_state['files'] = {}
            app_state['file_data'] = []
            app_state['filter_file'] = None
            self.send_json({'status': 'ok'})
        else:
            self.send_json({'error': 'Not found'}, 404)
    
    def handle_upload(self, content_length):
        """Handle file upload"""
        content_type = self.headers.get('Content-Type', '')
        
        if 'multipart/form-data' in content_type:
            # Parse multipart form data
            boundary = content_type.split('boundary=')[1].encode()
            body = self.rfile.read(content_length)
            
            parts = body.split(b'--' + boundary)
            files_processed = []
            
            for part in parts:
                if b'filename="' in part:
                    # Extract filename
                    header_end = part.find(b'\r\n\r\n')
                    header = part[:header_end].decode('utf-8', errors='ignore')
                    filename_start = header.find('filename="') + 10
                    filename_end = header.find('"', filename_start)
                    filename = header[filename_start:filename_end]
                    
                    # Extract file content
                    file_content = part[header_end + 4:]
                    if file_content.endswith(b'\r\n'):
                        file_content = file_content[:-2]
                    
                    if filename and file_content:
                        app_state['files'][filename] = file_content
                        files_processed.append(filename)
            
            # Process all uploaded files
            log_debug("=== FILE UPLOAD ===")
            log_debug(f"Files to process: {list(app_state['files'].keys())}")
            
            app_state['file_data'] = []
            for filename, content in app_state['files'].items():
                try:
                    file_info = self.process_file(filename, content)
                    app_state['file_data'].append(file_info)
                    log_debug(f"Processed '{filename}':")
                    for sheet in file_info.get('sheets', []):
                        log_debug(f"  Sheet '{sheet['name']}': {len(sheet.get('data', []))} rows, headers: {sheet.get('headers', [])}")
                except Exception as e:
                    log_error(f"Error processing {filename}", e)
                    self.send_json({'error': f'Error processing {filename}: {str(e)}'}, 400)
                    return
            
            self.send_json({'status': 'ok', 'files': app_state['file_data']})
        else:
            self.send_json({'error': 'Invalid content type'}, 400)
    
    def process_file(self, filename, content):
        """Process an Excel or CSV file"""
        file_info = {
            'fileName': filename,
            'fileType': 'csv' if filename.endswith('.csv') else 'excel',
            'sheets': []
        }
        
        def trim_value(val):
            """Trim whitespace from values"""
            return str(val).strip() if val is not None else ''
        
        try:
            if filename.endswith('.csv'):
                df = pd.read_csv(BytesIO(content))
                # Trim column headers
                df.columns = [str(c).strip() for c in df.columns]
                headers = df.columns.tolist()
                # Convert to strings and trim all values
                data = []
                for _, row in df.fillna('').iterrows():
                    data.append({col: trim_value(row[col]) for col in headers})
                file_info['sheets'].append({
                    'name': 'Sheet1',
                    'headers': headers,
                    'data': data
                })
            else:
                xlsx = pd.ExcelFile(BytesIO(content))
                for sheet_name in xlsx.sheet_names:
                    df = pd.read_excel(xlsx, sheet_name=sheet_name)
                    # Trim column headers
                    df.columns = [str(c).strip() for c in df.columns]
                    headers = df.columns.tolist()
                    # Convert to strings and trim all values
                    data = []
                    for _, row in df.fillna('').iterrows():
                        data.append({col: trim_value(row[col]) for col in headers})
                    file_info['sheets'].append({
                        'name': sheet_name,
                        'headers': headers,
                        'data': data
                    })
        except Exception as e:
            raise Exception(f'Failed to read file: {str(e)}')
        
        return file_info
    
    def handle_upload_filter(self, content_length):
        """Handle filter file upload"""
        content_type = self.headers.get('Content-Type', '')
        
        if 'multipart/form-data' in content_type:
            boundary = content_type.split('boundary=')[1].encode()
            body = self.rfile.read(content_length)
            
            parts = body.split(b'--' + boundary)
            
            for part in parts:
                if b'filename="' in part:
                    header_end = part.find(b'\r\n\r\n')
                    header = part[:header_end].decode('utf-8', errors='ignore')
                    filename_start = header.find('filename="') + 10
                    filename_end = header.find('"', filename_start)
                    filename = header[filename_start:filename_end]
                    
                    file_content = part[header_end + 4:]
                    if file_content.endswith(b'\r\n'):
                        file_content = file_content[:-2]
                    
                    if filename and file_content:
                        try:
                            file_info = self.process_file(filename, file_content)
                            app_state['filter_file'] = file_info
                            self.send_json({'status': 'ok', 'file': file_info})
                            return
                        except Exception as e:
                            self.send_json({'error': f'Error processing filter file: {str(e)}'}, 400)
                            return
            
            self.send_json({'error': 'No file found'}, 400)
        else:
            self.send_json({'error': 'Invalid content type'}, 400)
    
    def handle_validate_filter(self, content_length):
        """Validate filter matches - returns count of matching rows per source"""
        body = self.rfile.read(content_length)
        config = json.loads(body.decode('utf-8'))
        
        filter_values = config.get('filterValues', [])
        column_mappings = config.get('columnMappings', {})
        file_data = config.get('fileData', [])
        
        # Trim all filter values
        filter_values_set = set(str(v).strip().lower() for v in filter_values if str(v).strip())
        
        results = {}
        
        for source_key, filter_column in column_mappings.items():
            if not filter_column:
                continue
            
            parts = source_key.split('-')
            file_idx = int(parts[0])
            sheet_name = '-'.join(parts[1:])
            
            if file_idx >= len(file_data):
                continue
            
            file = file_data[file_idx]
            sheet = next((s for s in file['sheets'] if s['name'] == sheet_name), None)
            if not sheet:
                continue
            
            total_rows = len(sheet['data'])
            matching_rows = 0
            
            for row in sheet['data']:
                val = str(row.get(filter_column, '')).strip().lower()
                if val and val in filter_values_set:
                    matching_rows += 1
            
            results[source_key] = {
                'total': total_rows,
                'matching': matching_rows,
                'filterColumn': filter_column
            }
        
        self.send_json({'results': results})
    
    def handle_process(self, content_length):
        """Return processed file data"""
        self.send_json({'files': app_state['file_data']})
    
    def handle_compute(self, content_length):
        """Compute matrices based on configuration"""
        log_debug("=== COMPUTE MATRICES REQUEST ===")
        
        body = self.rfile.read(content_length)
        config = json.loads(body.decode('utf-8'))
        
        log_debug("Received config keys", list(config.keys()))
        log_debug("fileData count", len(config.get('fileData', [])))
        log_debug("selectedTabs", config.get('selectedTabs'))
        log_debug("columnSelections", config.get('columnSelections'))
        log_debug("matrixConfig", config.get('matrixConfig'))
        
        # Extract filter config if present
        filter_config = config.get('filterConfig')
        filter_values = None
        filter_mappings = None
        
        log_debug("filterConfig", filter_config)
        
        if filter_config and filter_config.get('enabled') and filter_config.get('values'):
            # Trim all filter values and convert to lowercase for case-insensitive matching
            filter_values = set(str(v).strip().lower() for v in filter_config['values'] if str(v).strip())
            filter_mappings = filter_config.get('columnMappings', {})
            log_debug(f"Filter enabled with {len(filter_values)} values")
            log_debug("Filter mappings", filter_mappings)
        else:
            log_debug("Filter NOT enabled or no values")
        
        try:
            matrices = self.compute_matrices(
                config['fileData'],
                config['selectedTabs'],
                config['columnSelections'],
                config['matrixConfig'],
                filter_values,
                filter_mappings
            )
            log_debug(f"Computed {len(matrices)} matrices")
            for m in matrices:
                log_debug(f"  Matrix '{m['name']}': {len(m['rows'])} rows x {len(m['cols'])} cols")
            
            self.send_json({'matrices': matrices})
        except Exception as e:
            log_error("Error in compute_matrices", e)
            self.send_json({'error': str(e)}, 400)
    
    def compute_matrices(self, file_data, selected_tabs, column_selections, matrix_config, filter_values=None, filter_mappings=None):
        """Compute intersection matrices
        
        New column_selections format:
            { key: { rowColumns: [...], colColumn: '...' } }
        
        - rowColumns: array of column names that combine to form row labels
        - colColumn: single column name for column values
        - filter_values: Optional set of values to filter rows by
        - filter_mappings: Optional dict { sourceKey: columnName } mapping each source to its filter column
        
        Output matrix:
        - rows: row labels (from rowColumns, combined)
        - cols: column headers (from colColumn)
        - data: 2D array [row_idx][col_idx] with 1s at intersections
        """
        log_debug("=== COMPUTE MATRICES START ===")
        log_debug(f"matrix_config count: {len(matrix_config)}")
        log_debug(f"file_data count: {len(file_data)}")
        log_debug(f"filter_values: {filter_values is not None} ({len(filter_values) if filter_values else 0} values)")
        
        matrices = []
        
        def get_row_value(row_data, row_columns):
            """Combine multiple columns into a single row label"""
            parts = []
            for col in row_columns:
                val = str(row_data.get(col, '')).strip()
                if val:
                    parts.append(val)
            return ' | '.join(parts) if parts else ''
        
        def get_filter_value(row_data, source_key, filter_mappings):
            """Get the value to use for filtering this row (lowercase for case-insensitive matching)"""
            if filter_mappings and source_key in filter_mappings:
                filter_col = filter_mappings[source_key]
                return str(row_data.get(filter_col, '')).strip().lower()
            return None  # No filter mapping for this source
        
        for config_idx, config in enumerate(matrix_config):
            log_debug(f"\n--- Processing matrix config {config_idx}: '{config.get('name', 'unnamed')}' ---")
            log_debug(f"Sources in config: {config.get('sources', [])}")
            log_debug(f"sourceKeys in config: {config.get('sourceKeys', [])}")
            
            # Collect all unique row and column values from all sources
            row_values = set()
            col_values = set()
            
            # Handle both 'sources' (array of objects) and 'sourceKeys' (array of strings) formats
            sources = config.get('sources', [])
            if not sources and config.get('sourceKeys'):
                # Convert sourceKeys to sources format
                log_debug("Converting sourceKeys to sources format")
                for sk in config['sourceKeys']:
                    parts = sk.split('-')
                    file_idx = int(parts[0])
                    sheet_name = '-'.join(parts[1:])
                    sources.append({'fileIndex': file_idx, 'sheetName': sheet_name})
                log_debug(f"Converted sources: {sources}")
            
            for source in sources:
                file_idx = source['fileIndex']
                sheet_name = source['sheetName']
                key = f"{file_idx}-{sheet_name}"
                
                log_debug(f"\n  Processing source: file_idx={file_idx}, sheet='{sheet_name}', key='{key}'")
                
                if file_idx >= len(file_data):
                    log_debug(f"  ERROR: file_idx {file_idx} >= file_data length {len(file_data)}")
                    continue
                
                file = file_data[file_idx]
                log_debug(f"  File: '{file.get('fileName', 'unknown')}'")
                log_debug(f"  Available sheets: {[s['name'] for s in file.get('sheets', [])]}")
                
                sheet = next((s for s in file['sheets'] if s['name'] == sheet_name), None)
                if not sheet:
                    log_debug(f"  ERROR: Sheet '{sheet_name}' not found in file")
                    continue
                
                log_debug(f"  Sheet data rows: {len(sheet.get('data', []))}")
                log_debug(f"  Sheet headers: {sheet.get('headers', [])}")
                
                sel = column_selections.get(key, {})
                log_debug(f"  Column selection for key '{key}': {sel}")
                
                row_columns = sel.get('rowColumns', [])
                col_column = sel.get('colColumn')
                
                log_debug(f"  rowColumns: {row_columns}")
                log_debug(f"  colColumn: {col_column}")
                
                if not row_columns or not col_column:
                    log_debug(f"  SKIPPING: Missing rowColumns or colColumn")
                    continue
                
                rows_processed = 0
                rows_added = 0
                cols_added = 0
                
                for row in sheet['data']:
                    rows_processed += 1
                    row_val = get_row_value(row, row_columns)
                    col_val = str(row.get(col_column, '')).strip()
                    
                    if row_val:
                        # Apply filter if enabled and this source has a filter mapping
                        if filter_values is not None and filter_mappings:
                            filter_val = get_filter_value(row, key, filter_mappings)
                            # If source has mapping, only include if value matches filter
                            if filter_val is not None:
                                if filter_val in filter_values:
                                    row_values.add(row_val)
                                    rows_added += 1
                            else:
                                # No mapping for this source, include all rows
                                row_values.add(row_val)
                                rows_added += 1
                        else:
                            # No filter enabled, include all rows
                            row_values.add(row_val)
                            rows_added += 1
                    if col_val:
                        col_values.add(col_val)
                        cols_added += 1
                
                log_debug(f"  Processed {rows_processed} rows, added {rows_added} unique row values, {cols_added} col values")
                if rows_processed > 0 and rows_added == 0:
                    # Sample some data for debugging
                    sample_rows = sheet['data'][:3]
                    log_debug(f"  SAMPLE DATA (first 3 rows):")
                    for sr in sample_rows:
                        log_debug(f"    Row: {sr}")
            
            sorted_rows = sorted(row_values)
            sorted_cols = sorted(col_values)
            
            log_debug(f"\n  Total unique rows collected: {len(sorted_rows)}")
            log_debug(f"  Total unique cols collected: {len(sorted_cols)}")
            
            if not sorted_rows or not sorted_cols:
                log_debug(f"  SKIPPING MATRIX: No rows ({len(sorted_rows)}) or no cols ({len(sorted_cols)})")
                continue
            
            # Create matrix data
            matrix_data = [[0] * len(sorted_cols) for _ in range(len(sorted_rows))]
            
            # Mark intersections
            for source in config['sources']:
                file_idx = source['fileIndex']
                sheet_name = source['sheetName']
                key = f"{file_idx}-{sheet_name}"
                
                file = file_data[file_idx]
                sheet = next((s for s in file['sheets'] if s['name'] == sheet_name), None)
                if not sheet:
                    continue
                
                sel = column_selections.get(key, {})
                row_columns = sel.get('rowColumns', [])
                col_column = sel.get('colColumn')
                
                if not row_columns or not col_column:
                    continue
                
                for row in sheet['data']:
                    row_val = get_row_value(row, row_columns)
                    col_val = str(row.get(col_column, '')).strip()
                    
                    # Apply same filter logic as above
                    should_include = True
                    if filter_values is not None and filter_mappings:
                        filter_val = get_filter_value(row, key, filter_mappings)
                        if filter_val is not None and filter_val not in filter_values:
                            should_include = False
                    
                    if should_include and row_val and col_val and row_val in sorted_rows and col_val in sorted_cols:
                        row_idx = sorted_rows.index(row_val)
                        col_idx = sorted_cols.index(col_val)
                        matrix_data[row_idx][col_idx] = 1
            
            matrices.append({
                'name': config['name'],
                'rows': sorted_rows,
                'cols': sorted_cols,
                'data': matrix_data
            })
        
        return matrices
    
    def handle_export(self, content_length):
        """Export matrices to Excel
        
        Matrix format:
        - rows: row labels (appear as first column, going down)
        - cols: column headers (appear as first row, going across)
        - data: 2D array [row_idx][col_idx]
        """
        log_debug("=== EXPORT REQUEST ===")
        
        body = self.rfile.read(content_length)
        data = json.loads(body.decode('utf-8'))
        matrices = data.get('matrices', [])
        
        log_debug(f"Matrices to export: {len(matrices)}")
        for m in matrices:
            log_debug(f"  '{m.get('name', 'unnamed')}': {len(m.get('rows', []))} rows x {len(m.get('cols', []))} cols")
        
        if not matrices:
            log_debug("ERROR: No matrices to export!")
            self.send_json({'error': 'No matrices to export. The matrices array is empty.'}, 400)
            return
        
        try:
            wb = Workbook()
            wb.remove(wb.active)
            
            # Define styles
            header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=12)
            title_font = Font(bold=True, size=14, color="1F4E79")
            label_font = Font(bold=True, size=11)
            normal_font = Font(size=11)
            thin_border = Border(
                left=Side(style='thin', color='D0D0D0'),
                right=Side(style='thin', color='D0D0D0'),
                top=Side(style='thin', color='D0D0D0'),
                bottom=Side(style='thin', color='D0D0D0')
            )
            center_align = Alignment(horizontal='center', vertical='center')
            left_align = Alignment(horizontal='left', vertical='center')
            dropdown_fill = PatternFill(start_color="E8F4FD", end_color="E8F4FD", fill_type="solid")
            result_fill = PatternFill(start_color="F0F9F0", end_color="F0F9F0", fill_type="solid")
            
            # ============================================
            # Create LOOKUP sheet as the first sheet
            # ============================================
            lookup_ws = wb.create_sheet(title="Lookup", index=0)
            
            # Collect all unique row values and build a mapping of row -> columns by matrix
            all_row_values = set()
            row_to_cols = {}  # { row_value: { matrix_name: [col1, col2, ...], ... } }
            
            for matrix in matrices:
                matrix_name = matrix['name']
                rows = matrix.get('rows', [])
                cols = matrix.get('cols', [])
                matrix_data = matrix.get('data', [])
                
                for row_idx, row_val in enumerate(rows):
                    all_row_values.add(row_val)
                    if row_val not in row_to_cols:
                        row_to_cols[row_val] = {}
                    
                    # Find all columns with value 1 for this row
                    matching_cols = []
                    if row_idx < len(matrix_data):
                        for col_idx, val in enumerate(matrix_data[row_idx]):
                            if val == 1 and col_idx < len(cols):
                                matching_cols.append(cols[col_idx])
                    
                    if matching_cols:
                        if matrix_name not in row_to_cols[row_val]:
                            row_to_cols[row_val][matrix_name] = []
                        row_to_cols[row_val][matrix_name].extend(matching_cols)
            
            sorted_row_values = sorted(all_row_values)
            
            # Title
            lookup_ws.cell(row=1, column=1, value="MATRIX LOOKUP")
            lookup_ws.cell(row=1, column=1).font = Font(bold=True, size=18, color="1F4E79")
            lookup_ws.merge_cells('A1:E1')
            
            lookup_ws.cell(row=2, column=1, value="Select a row value to see all matching columns across matrices")
            lookup_ws.cell(row=2, column=1).font = Font(size=11, italic=True, color="666666")
            lookup_ws.merge_cells('A2:E2')
            
            # Dropdown selection area
            lookup_ws.cell(row=4, column=1, value="Select Row Value:")
            lookup_ws.cell(row=4, column=1).font = label_font
            lookup_ws.cell(row=4, column=1).alignment = left_align
            
            # Cell for dropdown (B4)
            dropdown_cell = lookup_ws.cell(row=4, column=2, value=sorted_row_values[0] if sorted_row_values else "")
            dropdown_cell.fill = dropdown_fill
            dropdown_cell.font = normal_font
            dropdown_cell.border = thin_border
            dropdown_cell.alignment = left_align
            
            # Create data validation for dropdown
            if sorted_row_values:
                # Store row values in a hidden area (column H onwards) for the dropdown
                for idx, val in enumerate(sorted_row_values, start=1):
                    lookup_ws.cell(row=idx, column=8, value=val)
                
                # Hide column H
                lookup_ws.column_dimensions['H'].hidden = True
                
                # Create dropdown referencing the list
                dv = DataValidation(
                    type="list",
                    formula1=f"$H$1:$H${len(sorted_row_values)}",
                    allow_blank=False
                )
                dv.error = "Please select a value from the dropdown"
                dv.errorTitle = "Invalid Selection"
                dv.prompt = "Choose a row value"
                dv.promptTitle = "Row Selection"
                lookup_ws.add_data_validation(dv)
                dv.add(dropdown_cell)
            
            # Results header
            lookup_ws.cell(row=6, column=1, value="RESULTS")
            lookup_ws.cell(row=6, column=1).font = title_font
            
            lookup_ws.cell(row=7, column=1, value="Matrix")
            lookup_ws.cell(row=7, column=1).fill = header_fill
            lookup_ws.cell(row=7, column=1).font = header_font
            lookup_ws.cell(row=7, column=1).border = thin_border
            lookup_ws.cell(row=7, column=1).alignment = center_align
            
            lookup_ws.cell(row=7, column=2, value="Matching Columns")
            lookup_ws.cell(row=7, column=2).fill = header_fill
            lookup_ws.cell(row=7, column=2).font = header_font
            lookup_ws.cell(row=7, column=2).border = thin_border
            lookup_ws.cell(row=7, column=2).alignment = center_align
            lookup_ws.merge_cells('B7:E7')
            
            # Build a reference table for VLOOKUP (hidden in columns I, J)
            # Format: Composite Key (RowValue|||MatrixName) | Matching Columns
            # Using ||| as separator to avoid conflicts with data containing | or common chars
            ref_row = 1
            matrix_names = [m['name'] for m in matrices]
            
            for row_val in sorted_row_values:
                for matrix_name in matrix_names:
                    # Create composite key: RowValue|||MatrixName
                    composite_key = f"{row_val}|||{matrix_name}"
                    
                    # Get matching columns for this combination
                    if row_val in row_to_cols and matrix_name in row_to_cols[row_val]:
                        matching = ", ".join(sorted(set(row_to_cols[row_val][matrix_name])))
                    else:
                        matching = "(no matches)"
                    
                    lookup_ws.cell(row=ref_row, column=9, value=composite_key)
                    lookup_ws.cell(row=ref_row, column=10, value=matching)
                    ref_row += 1
            
            total_ref_rows = ref_row - 1
            
            # Hide reference columns
            lookup_ws.column_dimensions['I'].hidden = True
            lookup_ws.column_dimensions['J'].hidden = True
            
            # Add results area with formulas
            result_start_row = 8
            
            for idx, matrix_name in enumerate(matrix_names):
                current_row = result_start_row + idx
                
                # Matrix name cell
                lookup_ws.cell(row=current_row, column=1, value=matrix_name)
                lookup_ws.cell(row=current_row, column=1).font = label_font
                lookup_ws.cell(row=current_row, column=1).border = thin_border
                lookup_ws.cell(row=current_row, column=1).fill = result_fill
                lookup_ws.cell(row=current_row, column=1).alignment = left_align
                
                # Use simple VLOOKUP with composite key - this works in all Excel versions
                # Formula: =VLOOKUP(SelectedRow&"|||"&MatrixName, RefTable, 2, FALSE)
                # The composite key ensures exact match for both row value AND matrix name
                formula = f'=IFERROR(VLOOKUP($B$4&"|||"&A{current_row},$I$1:$J${total_ref_rows},2,FALSE),"(no matches)")'
                
                result_cell = lookup_ws.cell(row=current_row, column=2, value=formula)
                result_cell.font = normal_font
                result_cell.border = thin_border
                result_cell.fill = result_fill
                result_cell.alignment = left_align
                lookup_ws.merge_cells(f'B{current_row}:E{current_row}')
            
            # Set column widths for lookup sheet
            lookup_ws.column_dimensions['A'].width = 25
            lookup_ws.column_dimensions['B'].width = 50
            lookup_ws.column_dimensions['C'].width = 15
            lookup_ws.column_dimensions['D'].width = 15
            lookup_ws.column_dimensions['E'].width = 15
            
            # Add instructions
            instructions_row = result_start_row + len(matrix_names) + 2
            lookup_ws.cell(row=instructions_row, column=1, value="💡 Tip: Select a different value from the dropdown above to see matching columns for that row.")
            lookup_ws.cell(row=instructions_row, column=1).font = Font(size=10, italic=True, color="888888")
            lookup_ws.merge_cells(f'A{instructions_row}:E{instructions_row}')
            
            # ============================================
            # Create matrix sheets
            # ============================================
            for matrix in matrices:
                # Sanitize sheet name (max 31 chars, no special chars)
                sheet_name = matrix['name'][:31]
                for char in ['\\', '/', '*', '?', ':', '[', ']']:
                    sheet_name = sheet_name.replace(char, '_')
                
                ws = wb.create_sheet(title=sheet_name)
                
                rows = matrix.get('rows', [])
                cols = matrix.get('cols', [])
                matrix_data = matrix.get('data', [])
                
                # Header row: empty cell + column headers (styled)
                ws.cell(row=1, column=1, value='')
                ws.cell(row=1, column=1).fill = header_fill
                ws.cell(row=1, column=1).border = thin_border
                
                for col_idx, col_val in enumerate(cols, start=2):
                    cell = ws.cell(row=1, column=col_idx, value=col_val)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.border = thin_border
                    cell.alignment = center_align
                
                # Data rows: row label + values
                for row_idx, row_val in enumerate(rows, start=2):
                    # Row label
                    label_cell = ws.cell(row=row_idx, column=1, value=row_val)
                    label_cell.font = label_font
                    label_cell.border = thin_border
                    label_cell.fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
                    
                    # Data cells
                    if row_idx - 2 < len(matrix_data):
                        for col_idx, val in enumerate(matrix_data[row_idx - 2], start=2):
                            cell = ws.cell(row=row_idx, column=col_idx, value=val)
                            cell.border = thin_border
                            cell.alignment = center_align
                            if val == 1:
                                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                                cell.font = Font(bold=True, color="006100")
                
                # Auto-fit column widths (approximate)
                ws.column_dimensions['A'].width = max(12, max((len(str(r)) for r in rows), default=12) * 1.2)
                for col_idx in range(2, len(cols) + 2):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 12
            
            # Save to bytes
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'matrices_{timestamp}.xlsx'
            
            self.send_file_download(output.read(), filename)
        except Exception as e:
            self.send_json({'error': str(e)}, 500)


def run_server(port=8080):
    """Start the web server"""
    # Clear log file at startup
    clear_log()
    log_debug("Server starting...")
    
    server = HTTPServer(('127.0.0.1', port), MatrixProcessorHandler)
    print(f"\n{'='*50}")
    print(f"  MATRIX PROCESSOR")
    print(f"{'='*50}")
    print(f"\n  Server running at: http://localhost:{port}")
    print(f"\n  Debug log: {LOG_FILE}")
    print(f"\n  Opening browser...")
    print(f"\n  Keep this window open while using the app.")
    print(f"  Press Ctrl+C to stop.\n")
    print(f"{'='*50}\n")
    
    # Open browser after a short delay
    threading.Timer(1.0, lambda: webbrowser.open(f'http://localhost:{port}')).start()
    
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
        server.shutdown()


if __name__ == '__main__':
    run_server()

