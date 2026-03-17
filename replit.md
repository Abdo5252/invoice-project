# Invoice Data Extractor

A Streamlit web application that extracts structured invoice data from messy Excel files and outputs it in an organized format.

## Overview

The app processes Excel files containing multiple sheets with invoice data (supporting both Arabic and English headers), extracts key fields, and outputs a clean Excel file with two sheets: **Header** and **Items**.

## Architecture

- **Frontend/App**: Streamlit (Python) — single-page app running on port 5000
- **Language**: Python 3.12
- **Package Manager**: pip

## Key Files

- `app.py` — Main Streamlit application UI and orchestration
- `invoice_processor.py` — Core invoice parsing logic (reads Excel sheets, extracts fields)
- `excel_utils.py` — Utility functions for creating the output Excel file
- `requirements.txt` — Python dependencies
- `.streamlit/config.toml` — Streamlit server configuration (port 5000, host 0.0.0.0)

## Dependencies

- `streamlit` — Web app framework
- `pandas` — Data manipulation
- `numpy` — Numerical operations
- `openpyxl` — Excel file reading/writing

## Data Extraction

The app extracts:
- **Invoice Number**: Pattern `INVOICE N:` with format `SIxxxxx`
- **Customer Code**: Near `partner code:` with format `Cxxxx`
- **Currency**: Standardized as EGP or USD
- **Product Details**: From tables with Arabic (`التسمية`, `الكمية`, `سعر الوحدة`) or English (`Description`, `Quantity`, `Unit price`) headers

## Running the App

The workflow "Start application" runs: `streamlit run app.py`

The app is accessible at port 5000.
