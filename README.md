# Excel Duplicate Checker Pro

A professional web-based tool for detecting and analyzing duplicate records in Excel files (.xlsx, .xls) and CSV files. Built with Flask, this application provides comprehensive duplicate detection with advanced filtering options, data quality analysis, and detailed reporting.

![Excel Duplicate Checker Pro](https://img.shields.io/badge/Python-Flask-blue?logo=flask)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.0.0-orange)

## ✨ Features

### 🔍 Advanced Duplicate Detection
- **Multiple Detection Modes**: Show all duplicates, keep first occurrence, or keep last occurrence
- **Column-Specific Analysis**: Select specific columns for duplicate comparison
- **Smart Data Normalization**: Handles whitespace, case sensitivity, and common data variations
- **Real-time Preview**: See duplicates highlighted as you select columns

### 📊 Comprehensive Data Analysis
- **Data Quality Metrics**: Total rows, columns, blank cells, memory usage
- **Column Statistics**: Data types, unique values, null counts, sample data
- **Interactive Visualizations**: Color-coded duplicate highlighting
- **Performance Optimized**: Handles large files up to 50MB efficiently

### 📋 Professional Reporting
- **Excel Reports**: Formatted output with conditional formatting
- **Multiple File Formats**: Original data, duplicates only, cleaned data, analysis summary
- **ZIP Package Downloads**: All reports bundled for easy sharing
- **Detailed Documentation**: Comprehensive analysis breakdown

### 🎨 Modern User Interface
- **Responsive Design**: Works on desktop, tablet, and mobile devices
- **Drag & Drop Upload**: Intuitive file handling
- **Progress Indicators**: Real-time processing feedback
- **Bootstrap 5**: Professional, accessible UI components

## 🚀 Quick Start

### Prerequisites
- Python 3.7 or higher
- pip (Python package manager)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/excel-duplicate-checker-pro.git
   cd excel-duplicate-checker-pro
   ```

2. **Create a virtual environment**
   ```bash
   python -m venv venv
   
   # On Windows
   venv\Scripts\activate
   
   # On macOS/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**
   ```bash
   python app.py
   ```

5. **Access the application**
   Open your browser and navigate to `http://localhost:8080`

## 📖 Usage Guide

### 1. Upload Your File
- **Supported Formats**: .xlsx, .xls, .csv
- **File Size Limit**: 50MB maximum
- **Methods**: Drag & drop or click to browse

### 2. Select Sheet (Excel files only)
- Choose which worksheet to analyze
- Preview available sheets before selection

### 3. Configure Duplicate Detection
- **Select Columns**: Choose which columns to compare for duplicates
- **Detection Type**:
  - *Show All Duplicates*: Highlights all duplicate records
  - *Keep First, Remove Others*: Shows only additional duplicates (keeps first occurrence)
  - *Keep Last, Remove Others*: Shows only additional duplicates (keeps last occurrence)

### 4. Analyze Results
- **Statistics Dashboard**: View key metrics and data quality indicators
- **Duplicate Preview**: See highlighted duplicate records in your data
- **Data Analysis**: Detailed column-by-column breakdown

### 5. Download Reports
- **Comprehensive Package**: ZIP file containing all analysis results
- **Multiple Formats**: Excel files with professional formatting
- **Ready to Share**: Professional reports for stakeholders

## 🛠️ Technical Details

### Technology Stack
- **Backend**: Flask 3.0.0 (Python web framework)
- **Data Processing**: Pandas 2.1.4 (Data manipulation and analysis)
- **Excel Handling**: OpenPyXL 3.1.2 (Excel file processing)
- **Frontend**: Bootstrap 5, Font Awesome, Vanilla JavaScript
- **File Security**: Werkzeug secure filename handling

### Key Components
- **Smart Data Normalization**: Handles common data inconsistencies
- **Memory Efficient**: Optimized for large dataset processing
- **Secure File Handling**: Validates file types and sizes
- **Session Management**: Maintains user state across requests
- **Error Handling**: Comprehensive error catching and user feedback

## 📁 Project Structure

```
excel-duplicate-checker-pro/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Main web interface
├── uploads/              # Temporary file storage (auto-created)
├── results/              # Generated reports (auto-created)
├── static/               # CSS, JS, images (if any)
└── README.md            # This file
```

## 🔧 Configuration

### Environment Variables
```bash
# Optional: Set custom secret key
export SECRET_KEY="your-secret-key-here"

# Optional: Set custom upload size limit (in bytes)
export MAX_FILE_SIZE=52428800  # 50MB
```

### Application Settings
- **Upload Folder**: `uploads/` (auto-created)
- **Results Folder**: `results/` (auto-created)
- **Allowed Extensions**: `.xlsx`, `.xls`, `.csv`
- **File Size Limit**: 50MB (configurable)

## 🚀 Deployment

### Local Development
```bash
python app.py
```
Application runs on `http://localhost:8080`

### Production Deployment
For production deployment, consider:
- Using a WSGI server like Gunicorn
- Setting up a reverse proxy with Nginx
- Configuring environment variables
- Setting up proper logging and monitoring

### Docker Support (Optional)
Create a `Dockerfile` for containerized deployment:
```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8080
CMD ["python", "app.py"]
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

### Common Issues
- **Large File Processing**: Files over 20MB may take longer to process
- **Memory Usage**: Very large files (40MB+) may require more system RAM
- **Browser Compatibility**: Modern browsers recommended (Chrome 80+, Firefox 75+, Safari 13+)

### Getting Help
- Open an [issue](https://github.com/yourusername/excel-duplicate-checker-pro/issues) on GitHub
- Check the documentation in the code comments
- Review the error messages in the application

## 🔮 Future Enhancements

- [ ] API endpoint for programmatic access
- [ ] Batch processing multiple files
- [ ] Advanced data validation rules
- [ ] Export to additional formats (JSON, XML)
- [ ] Scheduled duplicate monitoring
- [ ] Integration with cloud storage services
- [ ] Advanced statistical analysis features

## 📊 Screenshots

### Main Interface
<img width="1118" height="517" alt="image" src="https://github.com/user-attachments/assets/761955f1-cd51-47e7-9b44-6b980ca942f4" />


### Duplicate Detection
<img width="1056" height="532" alt="image" src="https://github.com/user-attachments/assets/59f78436-54ae-4e37-b9ed-3a1835100e42" />

---

**Built with ❤️ using Flask and Python**

*Star ⭐ this repository if you find it helpful!*
Add comprehensive README
