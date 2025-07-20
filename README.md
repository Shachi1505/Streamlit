# HR Letters - DocuWave 📄

A powerful Streamlit-based web application for generating personalized HR letters and agreements from Excel data using Word templates.

## 🚀 Features

### 🔐 Secure Authentication
- HR login system with username/password authentication
- Session management for secure access

### 📊 Excel Data Processing
- Upload Excel files (.xlsx format)
- Support for multiple sheets
- Automatic column normalization (lowercase, underscores)
- Intelligent column detection for ID and Name fields

### 📝 Word Template Processing
- Upload Word templates (.docx format)
- Automatic placeholder detection using `«placeholder»` syntax
- Support for both paragraph and table content replacement
- Date formatting for timestamp fields

### 🎯 Multiple Generation Modes
1. **Individual Employee Selection** - Generate agreement for a single employee
2. **Serial Number Range** - Generate agreements for a range of employees
3. **All Employees** - Bulk generation for entire dataset

### 📦 File Management
- Individual file downloads for single agreements
- ZIP file generation for multiple agreements
- Organized file naming: `{EmployeeID}_{EmployeeName}_{SheetName}.docx`

## 🛠️ Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Setup
1. Clone or download the project files
2. Install required dependencies:
```bash
pip install streamlit pandas python-docx
```

3. Ensure you have the following files in your project directory:
   - `app.py` - Main application file
   - `style.css` - Custom styling (optional)
   - `zinnia_logo.jpg` - Company logo

## 🚀 Usage

### Starting the Application
```bash
streamlit run app.py
```

### Login Credentials
- **Username:** HR001
- **Password:** zinnia@2025

### Step-by-Step Process

1. **Login** - Enter your HR credentials to access the application

2. **Upload Excel File** - Select your Excel file containing employee data
   - Choose the appropriate sheet from the dropdown
   - Ensure your Excel has columns that match the placeholders in your Word template

3. **Upload Word Template** - Select your Word template file
   - Use `«placeholder»` syntax for dynamic content
   - Example: `«Employee Name»`, `«Employee ID»`, `«Date»`

4. **Select Generation Mode**:
   - **Individual**: Select a specific employee from the dropdown
   - **Range**: Specify start and end row numbers
   - **All**: Generate for all employees in the dataset

5. **Generate & Download** - Click the generate button and download your files

## 📋 Excel File Requirements

### Required Columns
Your Excel file must contain columns that match the placeholders in your Word template. The application automatically:
- Converts column names to lowercase
- Replaces spaces and hyphens with underscores
- Detects ID and Name columns automatically

### Example Excel Structure
| Employee_ID | Employee_Name | Department | Joining_Date | Salary |
|-------------|---------------|------------|--------------|--------|
| EMP001      | John Doe      | IT         | 2024-01-15   | 50000  |
| EMP002      | Jane Smith    | HR         | 2024-02-01   | 45000  |

## 📄 Word Template Guidelines

### Placeholder Syntax
Use double angle brackets for placeholders:
- `«Employee Name»`
- `«Employee ID»`
- `«Department»`
- `«Joining Date»`

### Supported Content Types
- **Paragraphs**: Regular text content
- **Tables**: Tabular data with placeholders
- **Headers/Footers**: Document metadata

### Date Formatting
Date fields are automatically formatted as "Month Day, Year" (e.g., "January 15, 2024")

## 🔧 Configuration

### Logo Path
Update the logo path in `app.py` line 31:
```python
img_base64 = get_base64_image(
    r"path/to/your/logo.jpg"
)
```

### Authentication
Modify credentials in `app.py`:
```python
USERNAME = "your_username"
PASSWORD = "your_password"
```

## 📁 Project Structure
```
StreamlitProject/
├── app.py              # Main application file
├── style.css           # Custom styling
├── zinnia_logo.jpg     # Company logo
└── README.md           # This file
```

## 🎨 Customization

### Styling
Modify `style.css` to customize the application appearance:
- Colors and fonts
- Layout and spacing
- Button styles
- Form elements

### Features
The application can be extended with:
- Additional file formats
- Email integration
- Database connectivity
- Advanced authentication
- Audit logging

## 🔒 Security Features

- Session-based authentication
- Secure file handling with temporary directories
- Input validation and sanitization
- Error handling and user feedback

## 🐛 Troubleshooting

### Common Issues

1. **Missing Columns Error**
   - Ensure your Excel file contains all required columns
   - Check column names match template placeholders

2. **File Upload Issues**
   - Verify file formats (.xlsx for Excel, .docx for Word)
   - Check file size limits

3. **Template Processing Errors**
   - Ensure placeholders use correct syntax: `«placeholder»`
   - Check for special characters in placeholder names

### Error Messages
- **"Invalid credentials"** - Check username/password
- **"Missing columns"** - Verify Excel file structure
- **"File not found"** - Check file paths and permissions

## 📞 Support

For technical support or feature requests, please contact your IT department or system administrator.

## 📄 License

This application is developed for internal use by Zinnia. All rights reserved.

---

**Version:** 1.0  
**Last Updated:** 2025  
**Developed for:** Zinnia HR Department 