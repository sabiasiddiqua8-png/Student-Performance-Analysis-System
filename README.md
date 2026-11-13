# Student Performance Analysis System

A Python-based Student Performance Analysis System designed to manage student academic records, analyze performance, generate visualizations, and export reports.

## Project Overview

The Student Performance Analysis System is a menu-driven Python application that uses SQLite for data storage and provides tools for analyzing student academic performance.

It helps administrators manage student records and generate useful academic performance reports.

## Features

- Admin Login System
- Add Student Records
- Display All Student Records
- Search Students by Roll Number and Name
- Update Student Marks
- Delete Student Records
- Find Class Topper
- Class Performance Statistics
- Subject-Wise Performance Analysis
- Student Performance Rankings
- Student Performance Graphs
- Subject Performance Graphs
- Generate PDF Reports
- Export Records to CSV
- Export Records to Excel

## Technologies Used

- Python
- SQLite
- Matplotlib
- ReportLab
- OpenPyXL
- CSV

## Project Structure

```text
Student_Performance_Analysis/
│
├── student_performance.py
├── README.md
├── requirements.txt
├── .gitignore
│
└── reports/
    ├── student_performance_report.pdf
    ├── student_performance.csv
    └── student_performance.xlsx
```

*Note: The database, generated reports, and local backup files are excluded from GitHub.*

## Installation and Setup

### 1. Clone the repository

```bash
git clone https://github.com/sabiasiddiqua8-png/Student-Performance-Analysis-System.git
```

### 2. Navigate to the project folder

```bash
cd Student-Performance-Analysis-System
```

### 3. Create a virtual environment

```bash
python -m venv .venv
```

### 4. Activate the virtual environment

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

### 5. Install dependencies

```bash
pip install -r requirements.txt
```

### 6. Run the application

```bash
python student_performance.py
```

## Database

The application uses SQLite to store student records locally. The database is created automatically when the application runs.

## Important Note

The included admin login credentials are for demonstration purposes only. This project is intended for learning and portfolio demonstration, not production use with sensitive student information.

## Future Improvements

- Password hashing and secure authentication
- Web-based user interface
- Advanced data analytics
- Machine Learning-based student performance prediction

## Author

Developed as a Python portfolio project to practice database management, data analysis, visualization, and report generation.