# Radiologist Report Generator

A desktop application built with Python and tkinter for radiologists to create and save medical reports in Microsoft Word format.

## Features

- **Patient Information Entry**: Enter patient name, age, and test type
- **Report Generation**: Write detailed diagnosis and reports
- **Word Document Export**: Automatically creates .docx files with professional formatting
- **Offline Operation**: No internet connection required
- **Automatic File Management**: Saves reports to a dedicated Reports/ folder
- **Input Validation**: Ensures all required fields are completed
- **Modern UI**: Clean, professional interface

## Requirements

- Python 3.7 or higher
- tkinter (usually comes with Python)
- python-docx
- pyinstaller (for creating executable)

## Installation

1. **Clone or download the project files**

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   python radiologist_app.py
   ```

## Usage

1. **Launch the application** - The main window will appear with the radiologist's name and current date displayed.

2. **Enter Patient Information**:
   - **Patient Name**: Enter the full name of the patient
   - **Age**: Enter the patient's age (numbers only, 1-150)
   - **Test Type**: Select from the dropdown menu (X-Ray, CT Scan, MRI, etc.)

3. **Write the Report**:
   - Use the large text area to write the diagnosis and detailed report
   - The text area supports scrolling for longer reports

4. **Save the Report**:
   - Click the "Save Report" button
   - The application will validate all inputs
   - A .docx file will be created in the Reports/ folder
   - The filename format: `PatientName_YYYYMMDD_HHMMSS.docx`

5. **Clear Form** (Optional):
   - Click "Clear Form" to reset all fields for a new report

## File Structure

```
radiologist-app/
├── radiologist_app.py      # Main application file
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── Reports/               # Generated reports folder (created automatically)
└── build/                 # PyInstaller build folder (created when building exe)
```

## Creating Executable (Windows)

To create a standalone .exe file for Windows:

1. **Install PyInstaller** (if not already installed):
   ```bash
   pip install pyinstaller
   ```

2. **Build the executable**:
   ```bash
   pyinstaller --onefile --windowed --name "RadiologistReportGenerator" radiologist_app.py
   ```

3. **Find the executable**:
   - The .exe file will be created in the `dist/` folder
   - You can distribute this file to other Windows computers

## Generated Report Format

The generated Word document includes:

- **Header**: "Radiology Report" (centered)
- **Radiologist Information**: Name and current date/time
- **Patient Information**: Name, age, and test type
- **Diagnosis/Report**: The detailed medical report text

## Customization

### Changing Radiologist Name

Edit line 15 in `radiologist_app.py`:
```python
self.radiologist_name = "Dr. Your Name Here"
```

### Adding Test Types

Edit the `test_types` list around line 95 in `radiologist_app.py`:
```python
test_types = [
    "X-Ray",
    "CT Scan", 
    "MRI",
    "Ultrasound",
    "Mammography",
    "Bone Density",
    "Your Custom Test Type",
    "Other"
]
```

## Troubleshooting

### Common Issues

1. **"Module not found" error**:
   - Make sure you've installed the requirements: `pip install -r requirements.txt`

2. **Permission denied when saving**:
   - Ensure you have write permissions in the application directory
   - The Reports/ folder will be created automatically

3. **Word document won't open**:
   - Ensure you have Microsoft Word or a compatible application installed
   - The files are saved in .docx format (modern Word format)

### System Requirements

- **Windows**: Windows 7 or higher
- **macOS**: macOS 10.12 or higher  
- **Linux**: Most distributions with Python 3.7+
- **Memory**: 100MB RAM minimum
- **Storage**: 50MB free space

## Security Notes

- This application runs completely offline
- No patient data is transmitted or stored in the cloud
- All reports are saved locally on your computer
- Consider implementing additional security measures for HIPAA compliance in clinical environments

## License

This project is provided as-is for educational and clinical use. Please ensure compliance with your local medical data protection regulations.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Ensure all dependencies are properly installed
3. Verify Python version compatibility (3.7+) 