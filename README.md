# Secure PDF Viewer (SPV)

## Overview
Secure PDF Viewer (SPV) is a Python application designed for securely viewing PDF files. It supports multiple input methods including direct PDF file paths, SQL credentials, and base64-encoded data. The application ensures security and ease of access, making it suitable for various use cases where document protection and integrity are paramount.

## Features
1. **PDF Viewing**: Open and view PDF files.
2. **Database Integration**: Fetch and display PDF files stored in a SQL database.
3. **Base64 Support**: Open PDF files from base64-encoded data.
4. **Secure Connection**: Use secure connection strings to fetch files.
5. **Navigation**: Easy navigation through PDF pages with zoom capabilities.

## Usage Instructions

### 1. Receiving Connection String (cn) from VBA
When you receive a connection string from VBA, the parameters must be in the following order:
```
python pdf_viewer.py <connection_string> <file_title> <password>
```
**Example:**
```
python pdf_viewer.py "Provider=Microsoft.Aceess.OLEDB.10.0;Persist Security Info=False;Data Source=XXXXX\SQLEXPRESS;User ID=sa;Initial Catalog=XXXXX;Data Provider=SQLOLEDB.1" test.pdf XXXXX
```

### 2. Receiving Direct SQL Credentials as Arguments
When passing SQL credentials directly, use the following order:
```
python pdf_viewer.py <username> <password> <file_title> <server> <catalog>
```
**Example:**
```
python pdf_viewer.py XXXXX XXXXX test.pdf XXXXX\SQLEXPRESS XXXXX
```

### 3. No Arguments Provided
If no arguments are provided, the application will open a file chooser dialog to select a PDF file:
```
python pdf_viewer.py
```

### 4. Invalid File Type
If a non-PDF file is provided, the application will display a message indicating that the file type is not supported:
```
python pdf_viewer.py <invalid_file_type>
```
**Note:** Currently, non-PDF files receive a base64 treatment but display a message box stating "Filetype not supported yet, contact admin to software admin".

## Example Commands
1. **VBA Connection String Example:**
   ```
   python pdf_viewer.py "Provider=Microsoft.Aceess.OLEDB.10.0;Persist Security Info=False;Data Source=XXXXX\SQLEXPRESS;User ID=sa;Initial Catalog=XXXX;Data Provider=SQLOLEDB.1" test.pdf XXXXXX
   ```

2. **Direct SQL Credentials Example:**
   ```
   python pdf_viewer.py sa Mmggit22 test.pdf CHIWO\SQLEXPRESS Aircare
   ```

3. **Open File Chooser Example:**
   ```
   python pdf_viewer.py
   ```

4. **Invalid File Type Example:**
   ```
   python pdf_viewer.py not_a_pdf_file.txt
   ```

## Installation
Ensure you have the necessary Python packages installed:
```sh
pip ....

pyinstaller --noconsole --icon="D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\icon.ico" --onefile --windowed --version-file=version_info.txt pdf_viewer.py 

```

#