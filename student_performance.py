import csv
import getpass
import sqlite3
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# STUDENT PERFORMANCE ANALYSIS SYSTEM
# ==========================================

# SUBJECTS
subjects = ["Python", "Java", "DBMS", "Mathematics", "Operating Systems"]


# ==========================================
# 1. LOAD STUDENT RECORDS
# ==========================================

def load_students():

    connection = sqlite3.connect("students.db")
    cursor = connection.cursor()

    cursor.execute("SELECT * FROM students")

    records = cursor.fetchall()

    students = []

    for record in records:

        marks = [
            record[2],
            record[3],
            record[4],
            record[5],
            record[6]
        ]

        student = [
            record[0],
            record[1],
            marks,
            record[7],
            record[8],
            record[9],
            record[10]
        ]

        students.append(student)

    connection.close()

    return students


# ==========================================
# 2. SAVE STUDENT RECORDS
# ==========================================


def insert_student(student):

    connection = sqlite3.connect("students.db")

    try:

        cursor = connection.cursor()

        cursor.execute("""
            INSERT INTO students VALUES
            (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            student[0],
            student[1],
            student[2][0],
            student[2][1],
            student[2][2],
            student[2][3],
            student[2][4],
            student[3],
            student[4],
            student[5],
            student[6]
        ))

        connection.commit()
        return True

    except sqlite3.IntegrityError:

        print("Error: Roll number already exists!")
        return False

    finally:
        connection.close()

def update_student_in_database(student):

    connection = sqlite3.connect("students.db")

    try:

        cursor = connection.cursor()

        cursor.execute("""
            UPDATE students
            SET python = ?,
                java = ?,
                dbms = ?,
                mathematics = ?,
                operating_systems = ?,
                total = ?,
                average = ?,
                grade = ?,
                status = ?
            WHERE roll_number = ?
        """, (
            student[2][0],
            student[2][1],
            student[2][2],
            student[2][3],
            student[2][4],
            student[3],
            student[4],
            student[5],
            student[6],
            student[0]
        ))

        if cursor.rowcount == 0:
            print("No student record was updated.")

        connection.commit()

        return cursor.rowcount > 0

    finally:
        connection.close()

def delete_student_from_database(roll_number):

    connection = sqlite3.connect("students.db")
    cursor = connection.cursor()

    cursor.execute(
        "DELETE FROM students WHERE roll_number = ?",
        (roll_number,)
    )

    if cursor.rowcount == 0:
        print("No student record was deleted.")

    connection.commit()

    connection.close()

def search_student_in_database(roll_number):

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM students WHERE roll_number = ?",
            (roll_number,)
        )

        record = cursor.fetchone()

        if record is None:
            return None

        marks = [
            record[2],
            record[3],
            record[4],
            record[5],
            record[6]
        ]

        student = [
            record[0],
            record[1],
            marks,
            record[7],
            record[8],
            record[9],
            record[10]
        ]

        return student

    finally:
        connection.close()

def display_students_from_database():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("SELECT * FROM students")
        records = cursor.fetchall()

        if not records:
            print("\nNo student records available.")
            return

        print("\n========== ALL STUDENT RECORDS ==========")

        for record in records:

            print("\n--------------------------------")
            print("Roll Number:", record[0])
            print("Name:", record[1])
            print("\nSubject Marks:")

            for i in range(len(subjects)):
                print(subjects[i], ":", record[i + 2])

            print("\nTotal Marks:", record[7])
            print("Average:", round(record[8], 2))
            print("Grade:", record[9])
            print("Status:", record[10])

        print("\n=========================================")
        print("Total Students:", len(records))

    finally:
        connection.close()

def search_student_by_name_from_database():

    search_name = input("Enter student name to search: ").strip()

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute(
            "SELECT * FROM students WHERE name LIKE ?",
            ('%' + search_name + '%',)
        )

        records = cursor.fetchall()

        if not records:
            print("\nStudent not found!")
            return

        print("\n========== SEARCH RESULTS ==========")

        for record in records:

            print("\n------------------------------")
            print("Roll Number:", record[0])
            print("Name:", record[1])
            print("\nSubject Marks:")

            for i in range(len(subjects)):
                print(subjects[i], ":", record[i + 2])

            print("\nTotal Marks:", record[7])
            print("Average:", round(record[8], 2))
            print("Grade:", record[9])
            print("Status:", record[10])

        print("\n====================================")

    finally:
        connection.close()

def find_topper_from_database():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT * FROM students
            WHERE average = (
                SELECT MAX(average) FROM students
            )
        """)

        toppers = cursor.fetchall()

        if not toppers:
            print("\nNo student records available.")
            return

        print("\n========== CLASS TOPPER ==========")

        for student in toppers:

            print("\nRoll Number:", student[0])
            print("Name:", student[1])
            print("Total Marks:", student[7])
            print("Average:", round(student[8], 2))
            print("Grade:", student[9])
            print("Status:", student[10])

        print("\n==================================")

    finally:
        connection.close()

def create_database():

    connection = sqlite3.connect("students.db")

    cursor = connection.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            roll_number TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            python REAL,
            java REAL,
            dbms REAL,
            mathematics REAL,
            operating_systems REAL,
            total REAL,
            average REAL,
            grade TEXT,
            status TEXT
        )
    """)

    connection.commit()

    connection.close()

    print("Database created successfully!")

def display_rankings_from_database():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT * FROM students
            ORDER BY average DESC, total DESC, name ASC
        """)

        records = cursor.fetchall()

        if not records:
            print("\nNo student records available.")
            return

        print("\n========== STUDENT RANKINGS ==========")

        for rank, student in enumerate(records, start=1):

            print("\nRank:", rank)
            print("Roll Number:", student[0])
            print("Name:", student[1])
            print("Total Marks:", student[7])
            print("Average:", round(student[8], 2))
            print("Grade:", student[9])
            print("Status:", student[10])

            print("--------------------------------------")

        print("Total Students:", len(records))
        print("======================================")

    finally:
        connection.close()

def class_statistics_from_database():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT
                COUNT(*),
                AVG(average),
                MAX(average),
                MIN(average),
                SUM(CASE WHEN status = 'Pass' THEN 1 ELSE 0 END),
                SUM(CASE WHEN status = 'Fail' THEN 1 ELSE 0 END)
            FROM students
        """)

        result = cursor.fetchone()

        total_students = result[0]

        if total_students == 0:
            print("\nNo student records available.")
            return

        average_performance = result[1]
        highest_average = result[2]
        lowest_average = result[3]
        passed_students = result[4]
        failed_students = result[5]

        print("\n========== CLASS PERFORMANCE STATISTICS ==========")

        print("Total Students:", total_students)
        print("Class Average:", round(average_performance, 2))
        print("Highest Average:", round(highest_average, 2))
        print("Lowest Average:", round(lowest_average, 2))
        print("Students Passed:", passed_students)
        print("Students Failed:", failed_students)

        print("==================================================")

    finally:
        connection.close()

def subject_analysis_from_database():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT
                AVG(python),
                AVG(java),
                AVG(dbms),
                AVG(mathematics),
                AVG(operating_systems)
            FROM students
        """)

        result = cursor.fetchone()

        if result[0] is None:
            print("\nNo student records available.")
            return

        print("\n========== SUBJECT-WISE PERFORMANCE ==========")

        for i in range(len(subjects)):

            print(
                subjects[i],
                "Average:",
                round(result[i], 2)
            )

        print("==============================================")

    finally:
        connection.close()

def check_database_records():

    connection = sqlite3.connect("students.db")
    cursor = connection.cursor()

    cursor.execute("SELECT COUNT(*) FROM students")

    count = cursor.fetchone()[0]

    print("\nTotal records in SQLite database:", count)

    cursor.execute("SELECT roll_number, name FROM students")

    records = cursor.fetchall()

    for record in records:
        print("Roll Number:", record[0], "| Name:", record[1])

    connection.close()

# Create SQLite database first
create_database()

# Load existing student records
students = load_students()


# ==========================================
# 3. ADD STUDENT DETAILS
# ==========================================

def add_students():

    try:
        number = int(input("Enter number of students to add: "))

        if number <= 0:
            print("Please enter a number greater than zero.")
            return

    except ValueError:
        print("Please enter a valid number.")
        return

    for i in range(number):

        print("\n----- STUDENT", i + 1, "DETAILS -----")

        roll_number = input("Enter roll number: ").strip()

        if any(student[0] == roll_number for student in students):
            print("This roll number already exists. Student not added.")
            continue

        name = input("Enter student name: ").strip()

        if not name:
            print("Name cannot be empty. Student not added.")
            continue

        marks = []

        valid_marks = True

        for subject in subjects:

            try:
                mark = float(input("Enter marks in " + subject + ": "))

                if mark < 0 or mark > 100:
                    print("Marks must be between 0 and 100.")
                    valid_marks = False
                    break

                marks.append(mark)

            except ValueError:
                print("Please enter valid numeric marks.")
                valid_marks = False
                break

        if not valid_marks:
            print("Student record was not added.")
            continue

        total = sum(marks)
        average = total / len(subjects)

        # Calculate grade

        if average >= 90:
            grade = "A+"

        elif average >= 80:
            grade = "A"

        elif average >= 70:
            grade = "B"

        elif average >= 60:
            grade = "C"

        elif average >= 50:
            grade = "D"

        else:
            grade = "F"

        # Calculate pass or fail

        if all(mark >= 40 for mark in marks):
            status = "Pass"

        else:
            status = "Fail"

        student = [
            roll_number,
            name,
            marks,
            total,
            average,
            grade,
            status
        ]

        if insert_student(student):

            students.append(student)

            print("Student added successfully!")

    print("\nStudent records saved successfully!")


# ==========================================
# 4. DISPLAY ALL STUDENTS
# ==========================================

def display_students():

    display_students_from_database()


# ==========================================
# 5. FIND CLASS TOPPER
# ==========================================

def find_topper():

    find_topper_from_database()


# ==========================================
# 6. SEARCH STUDENT BY ROLL NUMBER
# ==========================================

def search_student():

    search_roll = input("Enter roll number to search: ").strip()

    student = search_student_in_database(search_roll)

    if student is None:
        print("Student not found!")
        return

    print("\nStudent Found!")
    print("------------------------------")

    print("Roll Number:", student[0])
    print("Name:", student[1])

    print("\nSubject Marks:")

    for i in range(len(subjects)):
        print(subjects[i], ":", student[2][i])

    print("\nTotal Marks:", student[3])
    print("Average:", round(student[4], 2))
    print("Grade:", student[5])
    print("Status:", student[6])

def search_student_by_name():

    search_student_by_name_from_database()

# ==========================================
# 7. UPDATE STUDENT MARKS
# ==========================================

def update_student():

    if not students:
        print("No student records available.")
        return

    search_roll = input("Enter roll number to update: ").strip()

    for student in students:

        if student[0] == search_roll:

            print("\nStudent Found!")
            print("Name:", student[1])

            print("\nEnter the updated marks:")

            new_marks = []

            for subject in subjects:

                try:
                    mark = float(input("Enter new marks in " + subject + ": "))

                    if mark < 0 or mark > 100:
                        print("Marks must be between 0 and 100.")
                        print("Update cancelled.")
                        return

                    new_marks.append(mark)

                except ValueError:
                    print("Invalid marks entered. Update cancelled.")
                    return

            # Update marks
            student[2] = new_marks

            # Recalculate total and average
            student[3] = sum(new_marks)
            student[4] = student[3] / len(subjects)

            # Recalculate grade
            if student[4] >= 90:
                student[5] = "A+"
            elif student[4] >= 80:
                student[5] = "A"
            elif student[4] >= 70:
                student[5] = "B"
            elif student[4] >= 60:
                student[5] = "C"
            elif student[4] >= 50:
                student[5] = "D"
            else:
                student[5] = "F"

            # Recalculate pass or fail
            if all(mark >= 40 for mark in new_marks):
                student[6] = "Pass"
            else:
                student[6] = "Fail"

            if update_student_in_database(student):
                print("Student record updated successfully!")
            else:
                print("Error: Student record was not found in the database!")

            print("\nStudent marks updated successfully!")
            print("New Total:", student[3])
            print("New Average:", round(student[4], 2))
            print("New Grade:", student[5])
            print("New Status:", student[6])

            return

    print("Student not found!")


# ==========================================
# 8. DELETE STUDENT RECORD
# ==========================================

def delete_student():

    if not students:
        print("No student records available.")
        return

    search_roll = input("Enter roll number to delete: ").strip()

    for student in students:

        if student[0] == search_roll:

            print("\nStudent Found!")
            print("Name:", student[1])
            print("Roll Number:", student[0])

            confirm = input("Are you sure you want to delete this record? (yes/no): ")

            if confirm.lower() == "yes":

                delete_student_from_database(search_roll)

                students.remove(student)

                print("\nStudent record deleted successfully!")

            else:
                print("\nDeletion cancelled.")

            return

    print("Student not found!")


# ==========================================
# 9. CLASS PERFORMANCE STATISTICS
# ==========================================

def class_statistics():

    class_statistics_from_database()

# ==========================================
# 10. SUBJECT-WISE PERFORMANCE ANALYSIS
# ==========================================

def subject_analysis():

    subject_analysis_from_database()


# ==========================================
# 11. STUDENT PERFORMANCE GRAPH
# ==========================================

def student_performance_graph():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT name, average
            FROM students
            ORDER BY average DESC
        """)

        records = cursor.fetchall()

        if not records:
            print("\nNo student records available.")
            return

        names = [record[0] for record in records]
        averages = [record[1] for record in records]

        plt.figure(figsize=(10, 5))

        plt.bar(names, averages)

        plt.xlabel("Student Names")
        plt.ylabel("Average Marks")
        plt.title("Student Performance Analysis")

        plt.xticks(rotation=45)
        plt.tight_layout()

        plt.show()

    finally:
        connection.close()


# ==========================================
# 12. GENERATE PDF REPORT
# ==========================================

def generate_pdf_report():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("SELECT * FROM students ORDER BY roll_number")
        records = cursor.fetchall()

    finally:
        connection.close()

    if not records:
        print("\nNo student records available.")
        return

    file_name = "reports/student_performance_report.pdf"

    pdf = SimpleDocTemplate(
        file_name,
        pagesize=landscape(A4)
    )

    styles = getSampleStyleSheet()

    elements = []

    elements.append(
        Paragraph("Student Performance Analysis Report", styles["Title"])
    )

    elements.append(Spacer(1, 20))

    data = [[
        "Roll No.",
        "Name",
        "Python",
        "Java",
        "DBMS",
        "Maths",
        "OS",
        "Total",
        "Average",
        "Grade",
        "Status"
    ]]

    for student in records:

        data.append([
            student[0],
            student[1],
            student[2],
            student[3],
            student[4],
            student[5],
            student[6],
            student[7],
            round(student[8], 2),
            student[9],
            student[10]
        ])

    table = Table(data, repeatRows=1)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))

    elements.append(table)

    pdf.build(elements)

    print("\nPDF Report Generated Successfully!")
    print("File:", file_name)


# ==========================================
# 13. EXPORT STUDENT RECORDS TO CSV
# ==========================================

def export_to_csv():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT * FROM students
            ORDER BY roll_number
        """)

        records = cursor.fetchall()

    finally:
        connection.close()

    if not records:
        print("\nNo student records available.")
        return

    filename = "reports/student_performance.csv"

    try:
        with open(filename, "w", newline="", encoding="utf-8") as file:

            writer = csv.writer(file)

            header = [
                "Roll Number",
                "Student Name"
            ]

            header.extend(subjects)

            header.extend([
                "Total Marks",
                "Average",
                "Grade",
                "Status"
            ])

            writer.writerow(header)

            for student in records:

                row = list(student)

                row[8] = round(row[8], 2)

                writer.writerow(row)

        print("\nStudent records exported successfully!")
        print("File name:", filename)

    except OSError as error:
        print("Unable to create CSV file:", error)

def export_to_excel():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT * FROM students
            ORDER BY roll_number
        """)

        records = cursor.fetchall()

    finally:
        connection.close()

    if not records:
        print("\nNo student records available.")
        return

    try:

        workbook = Workbook()
        sheet = workbook.active

        sheet.title = "Student Performance"

        headers = [
            "Roll Number",
            "Student Name"
        ]

        headers.extend(subjects)

        headers.extend([
            "Total Marks",
            "Average",
            "Grade",
            "Status"
        ])

        sheet.append(headers)

        # Add student records
        for student in records:

            row = list(student)

            row[8] = round(row[8], 2)

            sheet.append(row)

        # Create Class Summary worksheet
        summary = workbook.create_sheet("Class Summary")

        summary["A1"] = "CLASS PERFORMANCE SUMMARY"

        summary["A3"] = "Total Students"
        summary["B3"] = len(records)

        summary["A4"] = "Class Average"
        summary["B4"] = round(
        sum(student[8] for student in records) / len(records), 2
        )

        summary["A5"] = "Highest Average"
        summary["B5"] = max(student[8] for student in records)

        summary["A6"] = "Lowest Average"
        summary["B6"] = min(student[8] for student in records)

        summary["A7"] = "Students Passed"
        summary["B7"] = sum(
            1 for student in records if student[10] == "Pass"
        )

        summary["A8"] = "Students Failed"
        summary["B8"] = sum(
            1 for student in records if student[10] == "Fail"
     )

        summary["A10"] = "Subject"
        summary["B10"] = "Average Marks"

        # Calculate subject-wise averages

        for i in range(len(subjects)):

            subject_average = sum(
                student[i + 2] for student in records
            ) / len(records)

            summary.cell(
                row=11 + i,
                column=1,
                value=subjects[i]
            )

            summary.cell(
                row=11 + i,
                column=2,
                value=round(subject_average, 2)
            )


        # Format summary headings
        summary["A1"].font = Font(size=16, bold=True, color="1F4E78")

        for cell in summary[10]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
           )

        summary.column_dimensions["A"].width = 25
        summary.column_dimensions["B"].width = 20

        # Format heading row
        for cell in sheet[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

        # Adjust column widths
        for column in sheet.columns:

            max_length = 0

            column_letter = get_column_letter(
                column[0].column
            )

            for cell in column:

                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            sheet.column_dimensions[column_letter].width = max_length + 4

        # Centre-align all cells
        for row in sheet.iter_rows():

            for cell in row:

                cell.alignment = Alignment(
                    horizontal="center"
                )

        # Freeze heading row
        sheet.freeze_panes = "A2"

        # Add Excel table formatting
        from openpyxl.worksheet.table import Table, TableStyleInfo

        table = Table(
            displayName="StudentPerformanceTable",
            ref=sheet.dimensions
        )

        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        table.tableStyleInfo = style

        sheet.add_table(table)

        # Add filters and improve display
        sheet.auto_filter.ref = sheet.dimensions
        sheet.sheet_view.showGridLines = False

        # Save workbook
        filename = "reports/student_performance.xlsx"

        workbook.save(filename)

        print("\nExcel file generated successfully!")
        print("File name:", filename)

    except Exception as error:

        print("Unable to generate Excel file:", error)

# ==========================================
# 14. SUBJECT PERFORMANCE GRAPH
# ==========================================

def subject_performance_graph():

    connection = sqlite3.connect("students.db")

    try:
        cursor = connection.cursor()

        cursor.execute("""
            SELECT
                AVG(python),
                AVG(java),
                AVG(dbms),
                AVG(mathematics),
                AVG(operating_systems)
            FROM students
        """)

        result = cursor.fetchone()

        if result[0] is None:
            print("\nNo student records available.")
            return

        averages = list(result)

        plt.figure(figsize=(10, 5))

        plt.bar(subjects, averages)

        plt.xlabel("Subjects")
        plt.ylabel("Average Marks")
        plt.title("Subject-Wise Performance Analysis")

        plt.xticks(rotation=30)
        plt.tight_layout()

        plt.show()

    finally:
        connection.close()


# ==========================================
# 15. DISPLAY STUDENT RANKINGS
# ==========================================

def display_rankings():

    display_rankings_from_database()

# ==========================================
# 16. ADMIN LOGIN SYSTEM
# ==========================================

def admin_login():

    correct_username = "admin"
    correct_password = "admin123"

    attempts = 3

    print("\n==============================================")
    print("       STUDENT PERFORMANCE SYSTEM LOGIN")
    print("==============================================")

    while attempts > 0:

        username = input("Enter admin username: ")

        password = getpass.getpass("Enter admin password: ")

        if username == correct_username and password == correct_password:

            print("\nLogin successful!")
            print("Welcome, Administrator!")

            return True

        else:

            attempts -= 1

            print("\nIncorrect username or password!")

            print("Attempts remaining:", attempts)

    print("\nAccess Denied!")
    print("Maximum login attempts reached.")

    return False

# Verify administrator login

if not admin_login():

    print("Program closed.")

    exit()

check_database_records()

# ==========================================
# 17. MAIN MENU
# ==========================================

while True:

    print("\n==============================================")
    print("       STUDENT PERFORMANCE ANALYSIS SYSTEM")
    print("==============================================")

    print("1. Add Student")
    print("2. Display All Students")
    print("3. Find Class Topper")
    print("4. Search Student")
    print("5. Search Student by Name")
    print("6. Update Student Marks")
    print("7. Delete Student Record")
    print("8. Class Performance Statistics")
    print("9. Subject-Wise Performance Analysis")
    print("10. Student Performance Graph")
    print("11. Subject Performance Graph")
    print("12. Generate PDF Report")
    print("13. Export Student Records to Excel (CSV)")
    print("14. Display Student Rankings")
    print("15. Export Student Records to Excel (.xlsx)")
    print("16. Exit")

    choice = input("\nEnter your choice (1-16): ")

    if choice == "1":

        add_students()

    elif choice == "2":

        display_students()

    elif choice == "3":

        find_topper()

    elif choice == "4":

        search_student()

    elif choice == "5":

        search_student_by_name()

    elif choice == "6":

        update_student()

    elif choice == "7":

        delete_student()

    elif choice == "8":

        class_statistics()

    elif choice == "9":

        subject_analysis()

    elif choice == "10":

        student_performance_graph()

    elif choice == "11":

        subject_performance_graph()

    elif choice == "12":

        generate_pdf_report()

    elif choice == "13":

        export_to_csv()

    elif choice == "14":

        display_rankings()

    elif choice == "15":

        export_to_excel()

    elif choice == "16":

        print("\nThank you for using our system!")
        print("Program closed successfully.")
        break

    else:

        print("Invalid choice! Please select between 1 and 16.")