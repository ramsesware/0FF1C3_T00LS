<div align="center">
    <img 
        src="https://github.com/ramsesware/ramsesware/blob/main/images/Office_Tools_Logo_Pharaoh.png"
        height=512
        weight=512
    />
</div>

---

```bash
   ██████╗ ███████╗███████╗██╗ ██████╗███████╗    ████████╗ ██████╗  ██████╗ ██╗     ███████╗
  ██╔═══██╗██╔════╝██╔════╝██║██╔════╝██╔════╝    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██╔════╝
  ██║   ██║█████╗  █████╗  ██║██║     █████╗         ██║   ██║   ██║██║   ██║██║     ███████╗
  ██║   ██║██╔══╝  ██╔══╝  ██║██║     ██╔══╝         ██║   ██║   ██║██║   ██║██║     ╚════██║
  ╚██████╔╝██║     ██║     ██║╚██████╗███████╗       ██║   ╚██████╔╝╚██████╔╝███████╗███████║
   ╚═════╝ ╚═╝     ╚═╝     ╚═╝ ╚═════╝╚══════╝       ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚══════╝
```                                                                                          

---


# OFFICE TOOLS

## Description

A collection of tools for office-related tasks. This repository will contain various utilities for document processing, file analysis, and data management in office environments.

## Table of Contents

- [About](#about)
- [Tools](#tools)
  - [ThotClean](#thotclean)
  - [Formatify](#formatify)
- [License](#license)

## About

**OFFICE TOOLS** is a suite of command-line tools for various office-related tasks, focusing on improving productivity and security in document handling. It aims to simplify tasks that often require manual and repetitive steps by providing automated solutions.

## Tools

### ThotClean

The first tool in the suite, **ThotClean**, is a metadata management utility for documents. ThotClean is designed to help analyze and remove metadata from various document formats, enhancing privacy and security. This tool is especially useful for users who want to ensure sensitive metadata is not unintentionally shared in files.

#### Features
- Analyzes metadata in supported document formats (PDF, DOCX, XLSX, PPTX, and image files).
- Removes unwanted metadata to ensure file privacy
- Supports multiple file formats commonly used in office environments

#### Dependencies
- `wxPython`
- `PyPDF2`
- `python-docx`
- `openpyxl`
- `python-pptx`
- `Pycryptodome`
- `hachoir`
- `mutagen`
- `Pillow`

#### How to install

```bash
pip install wxPython PyPDF2 python-docx openpyxl python-pptx Pycryptodome hachoir mutagen Pillow
```

```
If you prefer, you can download ThotClean.exe and you just have to install Python and execute
```
--- 

### Formatify

**Formatify** is a versatile file conversion tool that allows users to convert various document types (e.g., PDF, DOCX, XLSX, images) into other formats. It provides a GUI for selecting files and initiating conversions, making it user-friendly for office tasks.

#### Features
- Converts DOCX to PDF and plain text.
- Converts Excel (XLSX) to PDF and CSV.
- Extracts text from PDF.
- Converts PDF to images (JPEG, PNG).
- Converts images (JPEG, PNG) to PDF.
- Converts CSV to Excel (XLSX).
- Converts text files to PDF.

#### Dependencies
- `wxPython`
- `PyPDF2`
- `pdf2image`
- `FPDF`
- `python-docx`
- `pandas`
- `openpyxl`
- `Pillow`

#### How to install

```bash
pip install wxPython PyPDF2 pdf2image fpdf python-docx pandas openpyxl pillow
```



