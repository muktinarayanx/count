import http.server
import socketserver
import os
import re
import cgi
import pdfplumber
import openpyxl
from io import BytesIO

PORT = int(os.environ.get('PORT', 8080))

class ChallanHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/process':
            # Handle file uploads
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={'REQUEST_METHOD': 'POST',
                         'CONTENT_TYPE': self.headers['Content-Type'],
                         }
            )
            
            # Extract files
            pdf_files = []
            excel_template = None
            
            # Form fields can be list or single item
            for key in form.keys():
                fileitem = form[key]
                if isinstance(fileitem, list):
                    for item in fileitem:
                        if item.filename:
                            if item.filename.lower().endswith('.pdf'):
                                pdf_files.append((item.filename, item.file.read()))
                            elif item.filename.lower().endswith('.xlsx'):
                                excel_template = item.file.read()
                else:
                    if fileitem.filename:
                        if fileitem.filename.lower().endswith('.pdf'):
                            pdf_files.append((fileitem.filename, fileitem.file.read()))
                        elif fileitem.filename.lower().endswith('.xlsx'):
                            excel_template = fileitem.file.read()

            if not pdf_files:
                self.send_error(400, "No PDF files uploaded")
                return

            # Extract data from PDFs
            all_data = []
            for name, pdf_bytes in pdf_files:
                with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        if not text:
                            continue
                        
                        # 1. Challan extraction: Look for any 18-25 digit sequence
                        challan_match = re.search(r'\b\d{18,25}\b', text)
                        challan_no = challan_match.group(0) if challan_match else None
                        
                        # 2. Weight extraction: Look for all decimal/integer numbers followed by M t or Mt
                        weight = None
                        mt_matches = re.findall(r'(\d+\.?\d*)\s*[Mm]\s*[Tt]', text)
                        small_weights = [float(v) for v in mt_matches if float(v) < 1000 and float(v) > 0]
                        if small_weights:
                            weight = small_weights[0]
                        
                        if challan_no and weight is not None:
                            all_data.append({
                                'challan': challan_no,
                                'weight': weight,
                                'page': i + 1,
                                'fileName': name
                            })

            if not all_data:
                self.send_error(404, "No challan data found in the uploaded PDFs")
                return

            # Open template or create new workbook
            if excel_template:
                wb = openpyxl.load_workbook(BytesIO(excel_template))
            else:
                # Fallback to local file if exists
                if os.path.exists("challan check formula 1.xlsx"):
                    wb = openpyxl.load_workbook("challan check formula 1.xlsx")
                else:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Sheet1"
                    ws['A1'] = "unique no"
                    ws['B1'] = "Mt"

            ws = wb[wb.sheetnames[0]]
            start_row = 2

            # Write data in-place to preserve styling and conditional formatting rules
            for idx, d in enumerate(all_data):
                row = start_row + idx
                ws.cell(row=row, column=1, value=d['challan'])
                ws.cell(row=row, column=2, value=d['weight'])

            # Save workbook to memory
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            # Send response
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="challan_data_output.xlsx"')
            self.end_headers()
            self.wfile.write(output.read())
        else:
            super().do_POST()

# Start server
socketserver.TCPServer.allow_reuse_address = True
with socketserver.TCPServer(("", PORT), ChallanHandler) as httpd:
    print(f"Serving at port {PORT}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
