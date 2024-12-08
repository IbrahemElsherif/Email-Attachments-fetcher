# Outlook Email Attachment and PDF Processing Tool

## Overview

This Python script provides a comprehensive tool for:
- Fetching email attachments from Outlook based on a specific subject
- Processing PDF files with various functionalities

## Features

### Email Attachment Fetching
- Search emails by subject
- Extract attachments (all files or PDFs only)
- Save attachments to a specified directory

### PDF Processing
- Combine multiple PDF files
- Extract pages containing specific keywords
- Find word frequency across PDF documents

## Prerequisites

### Required Libraries
- `win32com.client`
- `pikepdf`
- `pdfplumber`

### Installation

1. Install required libraries:
```bash
pip install pywin32 pikepdf pdfplumber

~~~