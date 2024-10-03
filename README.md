# Bulk Email Sender

A Python script for sending bulk personalized emails using Yahoo's SMTP server. The script can read recipient details from CSV, Excel (.xlsx), or JSON files.

## Features

- Send personalized emails using a template with placeholders for names and surnames.
- Support for multiple input file formats: CSV, Excel, and JSON.
- Secure email sending using SSL encryption.

## Requirements

- Python 3.x
- Libraries:
  - `smtplib`
  - `email`
  - `ssl`
  - `csv`
  - `pandas` (for Excel file support)
  - `json`

You can install the required libraries using pip:

```bash
pip install pandas
