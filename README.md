DevTest - Brief Documentation

Project Overview:
DevTest is a Django-based web application designed to upload Excel/CSV files, process the data, generate summary reports, and send these reports via email. The application supports dynamic email content and generates visual representations of the data in image format for easier interpretation.

Features:
- File Upload: Users can upload Excel/CSV files.
- Data Processing: Processes uploaded files and generates summary reports.
- Dynamic Email Sending: Users can specify recipient addresses, email subjects, and bodies.
- Image Generation: Converts data into visual images with customizable formatting.
- Error Handling: Provides user-friendly error messages for invalid operations.

Requirements:
- Python 3.12
- Django 5.1
- Pillow (for image processing)
- openpyxl (for reading Excel files)
