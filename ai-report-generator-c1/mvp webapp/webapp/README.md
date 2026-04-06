# Container Inspection Report Management System

A web-based application for managing container inspection cases and generating professional reports.

## Features

✅ **User Authentication** - Secure login system with multiple user accounts
✅ **Document Upload** - Upload 4 types of documents (Bill of Lading, Commercial Invoice, Packing List, iAuditor/Safety Culture Reports)
✅ **Case Management** - Track and manage multiple inspection cases
✅ **Report Generation** - Automatically generate inspection reports in Word format
✅ **Dashboard** - Overview of all cases with statistics and status tracking

## Installation

### 1. Install Dependencies

```bash
cd webapp
pip install -r requirements.txt
```

### 2. Run the Application

```bash
python app.py
```

The application will start on `http://localhost:5001`

## Login Credentials

| Username  | Password   | Role      |
|-----------|------------|-----------|
| admin     | admin123   | Admin     |
| surveyor  | survey123  | Surveyor  |
| demo      | demo123    | Demo User |

## Usage Guide

### 1. Login
- Navigate to `http://localhost:5001`
- Enter username and password
- Click "Login"

### 2. Upload Documents
- After login, you'll be taken to the Upload page
- Fill in the case information:
  - Case Reference
  - Container Number
  - B/L Number
  - Goods Description
  - Shipper Name
  - Consignee Name
- Upload documents by clicking on each upload box
- Click "Create Case & Upload Documents"

### 3. Generate Report
- After uploading, you'll be redirected to the Generate Report page
- Review the case summary and uploaded documents
- Click "Generate Final Report"
- Download the generated report

### 4. Dashboard
- View all cases in the dashboard
- See statistics (Total, Pending, Completed)
- Download completed reports
- Generate reports for pending cases

## Folder Structure

```
webapp/
├── app.py                  # Flask backend application
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── templates/             # HTML templates
│   ├── login.html
│   ├── contact.html
│   ├── upload.html
│   ├── generate.html
│   └── dashboard.html
├── static/                # Static files
│   └── css/
│       └── style.css
├── uploads/               # Uploaded documents (created automatically)
├── reports/               # Generated reports (created automatically)
└── data/                  # Case data storage (created automatically)
```

## Customization

### Dashboard Layout
To customize the dashboard analytics section:
1. Open `templates/dashboard.html`
2. Find the section with class `custom-dashboard-section`
3. Replace the placeholder content with your custom layout

### Adding More Users
Edit the `USERS` dictionary in `app.py`:

```python
USERS = {
    'username': 'password',
    # Add more users here
}
```

### Styling
Modify `static/css/style.css` to customize the appearance.

## File Formats Supported

- **Documents**: PDF, DOC, DOCX
- **Images**: JPG, JPEG, PNG
- **Maximum file size**: 16MB

## Notes

- All data is stored in JSON format in the `data/` folder
- Reports are generated in Microsoft Word format (.docx)
- The system supports concurrent users
- Session-based authentication (sessions expire when browser closes)

## Security Notes

⚠️ **For Production Use:**
- Change the `app.secret_key` in `app.py`
- Use a proper database instead of JSON files
- Implement proper password hashing
- Use HTTPS
- Add user registration and password reset features

## Contact Admin

If you need help accessing your account:
- Email: admin@containerinspection.com
- Phone: +1 (555) 123-4567
- Office Hours: Monday - Friday, 9:00 AM - 5:00 PM

## Troubleshooting

### Port Already in Use
If port 5000 is already in use, change it in `app.py`:
```python
app.run(debug=True, host='0.0.0.0', port=5001)  # Change to 5001 or any available port
```

### Module Not Found Error
Make sure you've installed all dependencies:
```bash
pip install -r requirements.txt
```

### Permission Errors
Make sure the application has write permissions for:
- `uploads/` folder
- `reports/` folder
- `data/` folder

## License

This project is for internal use only.
