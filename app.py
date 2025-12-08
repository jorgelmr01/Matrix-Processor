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
from datetime import datetime
from io import BytesIO
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse

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
            app_state['file_data'] = []
            for filename, content in app_state['files'].items():
                try:
                    file_info = self.process_file(filename, content)
                    app_state['file_data'].append(file_info)
                except Exception as e:
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
        
        try:
            if filename.endswith('.csv'):
                df = pd.read_csv(BytesIO(content))
                headers = df.columns.tolist()
                data = df.fillna('').astype(str).to_dict('records')
                file_info['sheets'].append({
                    'name': 'Sheet1',
                    'headers': headers,
                    'data': data
                })
            else:
                xlsx = pd.ExcelFile(BytesIO(content))
                for sheet_name in xlsx.sheet_names:
                    df = pd.read_excel(xlsx, sheet_name=sheet_name)
                    headers = df.columns.tolist()
                    data = df.fillna('').astype(str).to_dict('records')
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
    
    def handle_process(self, content_length):
        """Return processed file data"""
        self.send_json({'files': app_state['file_data']})
    
    def handle_compute(self, content_length):
        """Compute matrices based on configuration"""
        body = self.rfile.read(content_length)
        config = json.loads(body.decode('utf-8'))
        
        # Extract filter config if present
        filter_config = config.get('filterConfig')
        filter_values = None
        if filter_config and filter_config.get('enabled') and filter_config.get('values'):
            filter_values = set(filter_config['values'])
        
        try:
            matrices = self.compute_matrices(
                config['fileData'],
                config['selectedTabs'],
                config['columnSelections'],
                config['matrixConfig'],
                filter_values
            )
            self.send_json({'matrices': matrices})
        except Exception as e:
            self.send_json({'error': str(e)}, 400)
    
    def compute_matrices(self, file_data, selected_tabs, column_selections, matrix_config, filter_values=None):
        """Compute intersection matrices
        
        Args:
            filter_values: Optional set of values to filter X axis (rows) by
        """
        matrices = []
        
        for config in matrix_config:
            if config.get('merge'):
                # Merge all sources into one matrix
                y_values = set()
                x_values = set()
                secondary_x_values = set()
                has_secondary_x = any(
                    column_selections.get(f"{s['fileIndex']}-{s['sheetName']}", {}).get('secondaryXAxis')
                    for s in config['sources']
                )
                
                # Collect unique values
                for source in config['sources']:
                    file_idx = source['fileIndex']
                    sheet_name = source['sheetName']
                    key = f"{file_idx}-{sheet_name}"
                    
                    file = file_data[file_idx]
                    sheet = next((s for s in file['sheets'] if s['name'] == sheet_name), None)
                    if not sheet:
                        continue
                    
                    sel = column_selections.get(key, {})
                    y_col = sel.get('yAxis')
                    x_col = sel.get('xAxis')
                    sec_x_col = sel.get('secondaryXAxis')
                    
                    for row in sheet['data']:
                        y_val = str(row.get(y_col, '')).strip()
                        x_val = str(row.get(x_col, '')).strip()
                        sec_x_val = str(row.get(sec_x_col, '')).strip() if sec_x_col else None
                        
                        if y_val:
                            y_values.add(y_val)
                        if x_val:
                            # Apply filter if present - only keep x values in the filter list
                            if filter_values is None or x_val in filter_values:
                                x_values.add(x_val)
                        if sec_x_val:
                            secondary_x_values.add(sec_x_val)
                
                sorted_y = sorted(y_values)
                sorted_x = sorted(x_values)
                sorted_sec_x = sorted(secondary_x_values) if has_secondary_x else None
                
                if sorted_sec_x:
                    for sec_x in sorted_sec_x:
                        matrix_data = [[0] * len(sorted_x) for _ in range(len(sorted_y))]
                        
                        for source in config['sources']:
                            file_idx = source['fileIndex']
                            sheet_name = source['sheetName']
                            key = f"{file_idx}-{sheet_name}"
                            
                            file = file_data[file_idx]
                            sheet = next((s for s in file['sheets'] if s['name'] == sheet_name), None)
                            if not sheet:
                                continue
                            
                            sel = column_selections.get(key, {})
                            y_col = sel.get('yAxis')
                            x_col = sel.get('xAxis')
                            sec_x_col = sel.get('secondaryXAxis')
                            
                            for row in sheet['data']:
                                y_val = str(row.get(y_col, '')).strip()
                                x_val = str(row.get(x_col, '')).strip()
                                sec_x_val = str(row.get(sec_x_col, '')).strip() if sec_x_col else None
                                
                                if y_val and x_val and sec_x_val == sec_x:
                                    y_idx = sorted_y.index(y_val) if y_val in sorted_y else -1
                                    x_idx = sorted_x.index(x_val) if x_val in sorted_x else -1
                                    if y_idx >= 0 and x_idx >= 0:
                                        matrix_data[y_idx][x_idx] = 1
                        
                        matrices.append({
                            'name': f"{config['name']} - {sec_x}",
                            'yAxis': sorted_y,
                            'xAxis': sorted_x,
                            'data': matrix_data
                        })
                else:
                    matrix_data = [[0] * len(sorted_x) for _ in range(len(sorted_y))]
                    
                    for source in config['sources']:
                        file_idx = source['fileIndex']
                        sheet_name = source['sheetName']
                        key = f"{file_idx}-{sheet_name}"
                        
                        file = file_data[file_idx]
                        sheet = next((s for s in file['sheets'] if s['name'] == sheet_name), None)
                        if not sheet:
                            continue
                        
                        sel = column_selections.get(key, {})
                        y_col = sel.get('yAxis')
                        x_col = sel.get('xAxis')
                        
                        for row in sheet['data']:
                            y_val = str(row.get(y_col, '')).strip()
                            x_val = str(row.get(x_col, '')).strip()
                            
                            if y_val and x_val:
                                y_idx = sorted_y.index(y_val) if y_val in sorted_y else -1
                                x_idx = sorted_x.index(x_val) if x_val in sorted_x else -1
                                if y_idx >= 0 and x_idx >= 0:
                                    matrix_data[y_idx][x_idx] = 1
                    
                    matrices.append({
                        'name': config['name'],
                        'yAxis': sorted_y,
                        'xAxis': sorted_x,
                        'data': matrix_data
                    })
            else:
                # Create independent matrix for each source
                for source in config['sources']:
                    file_idx = source['fileIndex']
                    sheet_name = source['sheetName']
                    key = f"{file_idx}-{sheet_name}"
                    
                    file = file_data[file_idx]
                    sheet = next((s for s in file['sheets'] if s['name'] == sheet_name), None)
                    if not sheet:
                        continue
                    
                    sel = column_selections.get(key, {})
                    y_col = sel.get('yAxis')
                    x_col = sel.get('xAxis')
                    sec_x_col = sel.get('secondaryXAxis')
                    
                    y_values = set()
                    x_values = set()
                    secondary_x_values = set() if sec_x_col else None
                    
                    for row in sheet['data']:
                        y_val = str(row.get(y_col, '')).strip()
                        x_val = str(row.get(x_col, '')).strip()
                        sec_x_val = str(row.get(sec_x_col, '')).strip() if sec_x_col else None
                        
                        if y_val:
                            y_values.add(y_val)
                        if x_val:
                            # Apply filter if present - only keep x values in the filter list
                            if filter_values is None or x_val in filter_values:
                                x_values.add(x_val)
                        if sec_x_val and secondary_x_values is not None:
                            secondary_x_values.add(sec_x_val)
                    
                    sorted_y = sorted(y_values)
                    sorted_x = sorted(x_values)
                    sorted_sec_x = sorted(secondary_x_values) if secondary_x_values else None
                    
                    if sorted_sec_x:
                        for sec_x in sorted_sec_x:
                            matrix_data = [[0] * len(sorted_x) for _ in range(len(sorted_y))]
                            
                            for row in sheet['data']:
                                y_val = str(row.get(y_col, '')).strip()
                                x_val = str(row.get(x_col, '')).strip()
                                sec_x_val = str(row.get(sec_x_col, '')).strip() if sec_x_col else None
                                
                                if y_val and x_val and sec_x_val == sec_x:
                                    y_idx = sorted_y.index(y_val) if y_val in sorted_y else -1
                                    x_idx = sorted_x.index(x_val) if x_val in sorted_x else -1
                                    if y_idx >= 0 and x_idx >= 0:
                                        matrix_data[y_idx][x_idx] = 1
                            
                            matrices.append({
                                'name': f"{source['fileName'].rsplit('.', 1)[0]} - {sheet_name} - {sec_x}",
                                'yAxis': sorted_y,
                                'xAxis': sorted_x,
                                'data': matrix_data
                            })
                    else:
                        matrix_data = [[0] * len(sorted_x) for _ in range(len(sorted_y))]
                        
                        for row in sheet['data']:
                            y_val = str(row.get(y_col, '')).strip()
                            x_val = str(row.get(x_col, '')).strip()
                            
                            if y_val and x_val:
                                y_idx = sorted_y.index(y_val) if y_val in sorted_y else -1
                                x_idx = sorted_x.index(x_val) if x_val in sorted_x else -1
                                if y_idx >= 0 and x_idx >= 0:
                                    matrix_data[y_idx][x_idx] = 1
                        
                        matrices.append({
                            'name': f"{source['fileName'].rsplit('.', 1)[0]} - {sheet_name}",
                            'yAxis': sorted_y,
                            'xAxis': sorted_x,
                            'data': matrix_data
                        })
        
        return matrices
    
    def handle_export(self, content_length):
        """Export matrices to Excel"""
        body = self.rfile.read(content_length)
        data = json.loads(body.decode('utf-8'))
        matrices = data.get('matrices', [])
        
        try:
            wb = Workbook()
            wb.remove(wb.active)
            
            for matrix in matrices:
                # Sanitize sheet name (max 31 chars, no special chars)
                sheet_name = matrix['name'][:31]
                for char in ['\\', '/', '*', '?', ':', '[', ']']:
                    sheet_name = sheet_name.replace(char, '_')
                
                ws = wb.create_sheet(title=sheet_name)
                
                # Header row
                ws.cell(row=1, column=1, value='')
                for col_idx, x_val in enumerate(matrix['xAxis'], start=2):
                    ws.cell(row=1, column=col_idx, value=x_val)
                
                # Data rows
                for row_idx, y_val in enumerate(matrix['yAxis'], start=2):
                    ws.cell(row=row_idx, column=1, value=y_val)
                    for col_idx, val in enumerate(matrix['data'][row_idx - 2], start=2):
                        ws.cell(row=row_idx, column=col_idx, value=val)
            
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
    server = HTTPServer(('127.0.0.1', port), MatrixProcessorHandler)
    print(f"\n{'='*50}")
    print(f"  MATRIX PROCESSOR")
    print(f"{'='*50}")
    print(f"\n  Server running at: http://localhost:{port}")
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
