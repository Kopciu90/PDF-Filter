# PDF Filter

## Overview

PDF Filter is a Python script designed to streamline the process of filtering and processing PDF documents based on specific criteria extracted from an Excel file. This tool reads account numbers (NRB) listed in an Excel spreadsheet and then searches through PDF documents to perform actions like removing pages containing these identifiers.

## Features

- Reads identifiers from an Excel spreadsheet.
- Searches through PDF documents for these identifiers.
- Removes pages from PDF documents based on the presence of identifiers.
- Supports batch processing of multiple PDFs.

## Prerequisites

- Python 3.x
- openpyxl (for handling Excel files)
- PyPDF2 (for PDF manipulation)
- pdfminer.six (for PDF text extraction)
