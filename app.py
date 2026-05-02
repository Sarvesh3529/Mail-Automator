import os
from flask import Flask, render_template, request, jsonify, Response, stream_with_context
from pathlib import Path
import tempfile
import file as email_logic

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send', methods=['POST'])
def send_emails():
    email = request.form.get('email')
    password = request.form.get('password')
    mapping_mode = request.form.get('mapping_mode', 'match')
    
    excel_file = request.files.get('excel')
    pdf_files = request.files.getlist('pdfs')

    if not all([email, password, excel_file, pdf_files]):
        return jsonify({"error": "Missing required fields"}), 400

    # We save files here before the response starts because Flask closes 
    # the request file streams immediately after this function returns.
    tmp_dir_obj = tempfile.TemporaryDirectory()
    tmp_path = Path(tmp_dir_obj.name)
    
    excel_path = tmp_path / "data.xlsx"
    excel_file.save(excel_path)
    
    saved_pdfs = []
    for pdf in pdf_files:
        p = tmp_path / pdf.filename
        pdf.save(p)
        saved_pdfs.append(p)

    def generate():
        try:
            yield "data: Reading your contact list...\n\n"
            yield f"data: Preparing {len(saved_pdfs)} documents...\n\n"

            try:
                rows = email_logic.read_xlsx_rows(excel_path)
                if not rows:
                    yield "data: Error: The Excel file is empty. Please check your data.\n\n"
                    return

                # Capture logic notes to display them on the web console
                notes = []
                jobs = email_logic.build_jobs(rows, saved_pdfs, mode=mapping_mode, logger=notes.append)
                for note in notes:
                    yield f"data: {note}\n\n"
                
                if not jobs:
                    yield "data: Error: No matching documents were found for your contacts.\n\n"
                    return

                yield f"data: Found {len(jobs)} matches! Sending emails now...\n\n"
                
                for status_msg in email_logic.send_with_gmail(jobs, email, password):
                    yield f"data: {status_msg}\n\n"
                
                yield "data: Success! All emails have been processed.\n\n"
            except Exception as e:
                err = str(e).lower()
                if "authentication" in err or "password" in err:
                    yield "data: Error: Login failed. Check your Gmail address and App Password.\n\n"
                elif "network" in err or "reach" in err:
                    yield "data: Error: Connection lost. Check your internet.\n\n"
                else:
                    yield f"data: Something went wrong: {str(e)}\n\n"
        finally:
            tmp_dir_obj.cleanup()

    return Response(stream_with_context(generate()), mimetype='text/event-stream')

if __name__ == '__main__':
    app.run(debug=True)