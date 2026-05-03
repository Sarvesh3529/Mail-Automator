import argparse
import os
import re
import smtplib
import ssl
import sys
import textwrap
import zipfile
import base64
from pathlib import Path
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET
from dotenv import load_dotenv
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Trial Mail.xlsx"
PDF_DIR = BASE_DIR / "PDF Files"
XML_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# Fetches from environment variables or uses placeholders
GMAIL_USER = os.getenv("GMAIL_USER", "your-email@gmail.com")
GMAIL_PASS = os.getenv("GMAIL_PASS", "")

# Change these if you want a different message.
DEFAULT_SUBJECT = "Pending GSTR-3B return"
DEFAULT_BODY = textwrap.dedent(
    """\
    Dear Taxpayer,

    Please file your pending GSTR-3B.

    Regards,
    XYZ
    """
)

def normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.strip().lower())

def read_xlsx_rows(xlsx_path: Path) -> List[Dict[str, str]]:
    def get_val(cell, strings):
        v = cell.find("main:v", XML_NS)
        if cell.get("t") == "s" and v is not None: return strings[int(v.text)]
        if cell.get("t") == "inlineStr":
            t = cell.find(".//main:t", XML_NS)
            return t.text.strip() if t is not None and t.text else ""
        return v.text.strip() if v is not None and v.text else ""

    with zipfile.ZipFile(xlsx_path) as workbook_zip:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in workbook_zip.namelist():
            shared_root = ET.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
            shared_strings = ["".join(t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")).strip() for si in shared_root.findall("main:si", XML_NS)]

        rels_root = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}
        
        wb_root = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
        relation_id = wb_root.find("main:sheets", XML_NS)[0].attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        
        sheet_target = "xl/" + rel_map[relation_id]
        sheet_root = ET.fromstring(workbook_zip.read(sheet_target))
        sheet_data = sheet_root.find("main:sheetData", XML_NS)

        rows: List[Dict[str, str]] = []
        headers: List[str] = []

        for row_index, row in enumerate(sheet_data.findall("main:row", XML_NS), start=1):
            values = [get_val(c, shared_strings) for c in row.findall("main:c", XML_NS)]
            if row_index == 1: headers = values
            elif any(values):
                rows.append({headers[i]: (values[i] if i < len(values) else "") for i in range(len(headers))})
        return rows

def extract_pdf_number(pdf_path: Path) -> Optional[str]:
    match = re.search(r"(\d+)\.pdf$", pdf_path.name, flags=re.IGNORECASE)
    return match.group(1) if match else None


def build_pdf_map(pdf_dir: Path) -> Dict[str, Path]:
    pdf_map: Dict[str, Path] = {}
    for pdf_path in sorted(pdf_dir.glob("*.pdf")):
        number = extract_pdf_number(pdf_path)
        if not number:
            print(f"Skipping PDF with unsupported name format: {pdf_path.name}")
            continue
        pdf_map[number] = pdf_path
    return pdf_map


def format_mail_body(record: Dict[str, str], body_template: str = None) -> str:
    norm = {normalize_header(k): v.strip() for k, v in record.items()}
    fields = {"Legal Name": "legalname", "Trade Name": "tradename", "GSTIN": "gstin"}
    details = [f"{label}: {norm[key]}" for label, key in fields.items() if norm.get(key)]
    
    template = body_template or DEFAULT_BODY.strip()
    if not details: return template
    
    prefix = f"{' | '.join(details)}\n\n"
    body = f"Dear Taxpayer,\n\n{prefix}{template}"
    return body

def build_jobs(rows: List[Dict[str, str]], pdf_files: List[Path], mode: str = "match", logger=print) -> List[Dict[str, str]]:
    if mode == "sequential":
        jobs = []
        sorted_pdfs = sorted(pdf_files)
        for i, record in enumerate(rows):
            if i >= len(sorted_pdfs):
                logger(f"Stopped at row {i+1}: No more PDFs found to attach.")
                break
            norm = {normalize_header(key): value.strip() for key, value in record.items()}
            recipient = norm.get("mail", "")
            if not recipient: continue
            jobs.append({"to": recipient, "subject": DEFAULT_SUBJECT, "body": format_mail_body(record), "attachment": str(sorted_pdfs[i])})
        return jobs

    pdf_map = {extract_pdf_number(p): p for p in pdf_files if extract_pdf_number(p)}
    jobs, missing_pdfs, missing_emails = [], [], []

    for record in rows:
        norm = {normalize_header(key): value.strip() for key, value in record.items()}
        sr_no, recipient = norm.get("srno", ""), norm.get("mail", "")

        if not sr_no: continue
        if not recipient:
            logger(f"Skipping Sr No {sr_no}: No email address found.")
            continue

        if sr_no not in pdf_map:
            logger(f"Skipping Sr No {sr_no}: Matching PDF not found.")
            continue

        jobs.append({
            "sr_no": sr_no,
            "to": recipient,
            "subject": DEFAULT_SUBJECT,
            "body": format_mail_body(record),
            "attachment": str(pdf_map[sr_no]),
        })

    return jobs


def send_with_gmail(jobs: List[Dict[str, str]], user: str, password: str):
    context = ssl.create_default_context()
    server = None
    
    # Prioritize Port 587 as confirmed working in this environment
    connection_configs = [
        (587, False), # Port 587: STARTTLS
        (465, True)   # Port 465: Explicit SSL (Fallback)
    ]
    
    last_exception = None

    try:
        for port, use_ssl in connection_configs:
            try:
                yield f"Attempting connection on Port {port}..."
                if use_ssl:
                    server = smtplib.SMTP_SSL("smtp.gmail.com", port, context=context, timeout=20)
                else:
                    server = smtplib.SMTP("smtp.gmail.com", port, timeout=20)
                    server.starttls(context=context)
                
                server.login(user, password)
                yield f"Connected successfully via Port {port}."
                break # Success, exit the loop
            except Exception as e:
                last_exception = e
                if server:
                    try: server.quit()
                    except: pass
                server = None
                yield f"Port {port} unreachable, trying fallback..."
                continue

        if not server:
            error_msg = str(last_exception)
            if "10060" in error_msg or "10061" in error_msg or "timed out" in error_msg.lower():
                raise RuntimeError(f"Network Blocked: Your firewall or ISP is blocking SMTP ports. (Details: {error_msg})")
            raise RuntimeError(f"Connection failed: {error_msg}")

        for job in jobs:
            msg = MIMEMultipart()
            msg['From'] = user
            msg['To'] = job['to']
            msg['Subject'] = job['subject']
            msg.attach(MIMEText(job['body'], 'plain'))

            attachment_path = Path(job['attachment'])
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{attachment_path.name}"')
            msg.attach(part)

            try:
                server.sendmail(user, job['to'], msg.as_string())
                yield f"Successfully sent to {job['to']}."
            except Exception as e:
                yield f"Failed for {job['to']}: {str(e)}"

    except Exception as e:
        raise RuntimeError(str(e))
    finally:
        if server:
            try: server.quit()
            except: pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send PDFs from 'PDF Files' to matching emails from 'Trial Mail.xlsx'."
    )
    return parser.parse_args()


def main() -> int:
    parse_args()
    if not EXCEL_FILE.exists() or not PDF_DIR.exists():
        print("Error: Missing Excel file or PDF directory.")
        return 1

    print(f"Reading Excel: {EXCEL_FILE.name}...")
    rows = read_xlsx_rows(EXCEL_FILE)
    print(f"Scanning PDFs: {PDF_DIR.name}...")
    pdf_files = list(PDF_DIR.glob("*.pdf"))
    jobs = build_jobs(rows, pdf_files)

    if not jobs:
        print("No valid jobs found. Check Excel data and PDF names.")
        return 0

    print(f"Prepared {len(jobs)} email jobs.")
    try:
        for msg in send_with_gmail(jobs, GMAIL_USER, GMAIL_PASS):
            print(f"  [Gmail] {msg}")
    except Exception as exc:
        print(f"\nError: {exc}")
        return 1

    print(f"\nCompleted successfully. Processed {len(jobs)} mail(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())