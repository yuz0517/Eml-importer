# Eml-importer
A lightweight Windows tool to bulk-import .eml files into Outlook PST folders, preserving sender, date, body, and attachments.
# EML Importer for Outlook

Batch import `.eml` files exported from groupware into Microsoft Outlook — 
no Python installation required.

## Why?

POP3-based mail sync doesn't carry over **Sent Mail**.  
This tool lets you download `.eml` files from your groupware and 
bulk-import them directly into any Outlook folder, preserving the 
original sender, date, body (HTML), and attachments.

## Download

Go to the **[Releases](../../releases)** tab and download `EML_Importer.exe`.  
No installation needed — just double-click and run.

## Usage

1. Open **Microsoft Outlook** first
2. Run `EML_Importer.exe`
3. Select the folder containing your `.eml` files
4. Choose the target Outlook folder (e.g. *Sent Items*)
5. Click **Start Import**

## Requirements

- Windows 10 / 11
- Microsoft Outlook (365, 2019, or 2021)

## What Gets Imported

| Field | Supported |
|---|---|
| Subject | ✓ |
| From / To / CC | ✓ |
| Original sent date | ✓ |
| HTML & plain text body | ✓ |
| Attachments | ✓ |
| Subfolders (recursive) | ✓ |
