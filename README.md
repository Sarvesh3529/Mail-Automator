# GST Email Dispatcher

A Python-based web application to automate sending PDF attachments to recipients listed in an Excel spreadsheet via Gmail.

## Features
- **Web Dashboard**: Simple interface to upload Excel files and PDFs.
- **Live Console**: Real-time status updates on email delivery.
- **Admin Quick-Login**: Easy access for authorized users.
- **Flexible Mapping**: Match PDFs by serial number or send them sequentially.

## Setup Instructions

1. **Install Dependencies**:
   ```bash
   pip install Flask
   ```

2. **Run the Application**:
   ```bash
   python app.py
   ```

3. **Usage**:
   - Open `http://127.0.0.1:5000` in your browser.
   - Use the **Admin** button (Password: `Dogesh`) to auto-fill credentials or enter your own Gmail App Password.