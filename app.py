
import sqlite3
import re
from io import BytesIO
from pathlib import Path
from datetime import datetime

from flask import (
    Flask, request, redirect, url_for,
    jsonify, render_template_string, send_file
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "attendance.db"


# ==================================================
# DATABASE
# ==================================================

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            prn TEXT NOT NULL UNIQUE,
            class_name TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            attendance_date TEXT NOT NULL,
            attendance_time TEXT NOT NULL,
            UNIQUE(student_id, attendance_date),
            FOREIGN KEY(student_id) REFERENCES students(id)
        )
    """)

    conn.commit()
    conn.close()


def valid_student(name, prn, class_name):
    return (
        bool(name.strip())
        and bool(class_name.strip())
        and bool(re.fullmatch(r"\d+", prn.strip()))
    )


# ==================================================
# DASHBOARD
# ==================================================

@app.route("/")
def home():
    conn = get_db()

    students = conn.execute("""
        SELECT * FROM students
        ORDER BY class_name, name
    """).fetchall()

    today = datetime.now().strftime("%Y-%m-%d")

    attendance = conn.execute("""
        SELECT
            students.name,
            students.prn,
            students.class_name,
            attendance.attendance_date,
            attendance.attendance_time
        FROM attendance
        JOIN students
            ON students.id = attendance.student_id
        WHERE attendance.attendance_date = ?
        ORDER BY attendance.attendance_time DESC
    """, (today,)).fetchall()

    conn.close()

    message = request.args.get("message", "")

    return render_template_string(
        PAGE,
        students=students,
        attendance=attendance,
        today=today,
        message=message
    )


# ==================================================
# SINGLE STUDENT REGISTRATION
# ==================================================

@app.route("/register", methods=["POST"])
def register():
    name = request.form.get("name", "").strip()
    prn = request.form.get("prn", "").strip()
    class_name = request.form.get("class_name", "").strip()

    if not valid_student(name, prn, class_name):
        return redirect(url_for(
            "home",
            message="Enter a valid name, numeric PRN and class."
        ))

    conn = get_db()

    try:
        conn.execute("""
            INSERT INTO students (name, prn, class_name)
            VALUES (?, ?, ?)
        """, (name, prn, class_name))

        conn.commit()
        message = "Student registered successfully!"

    except sqlite3.IntegrityError:
        message = "This PRN is already registered."

    finally:
        conn.close()

    return redirect(url_for("home", message=message))


# ==================================================
# BULK REGISTRATION - MANUAL ENTRY
# ==================================================

@app.route("/bulk-register", methods=["POST"])
def bulk_register():

    class_name = request.form.get("bulk_class", "").strip()

    names = request.form.getlist("student_name")
    prns = request.form.getlist("student_prn")

    if not class_name:
        return redirect(url_for(
            "home",
            message="Please enter the class name."
        ))

    if len(names) != len(prns):
        return redirect(url_for(
            "home",
            message="Student names and PRNs do not match."
        ))

    conn = get_db()

    added = 0
    duplicates = 0
    invalid = 0

    try:
        for name, prn in zip(names, prns):
            name = name.strip()
            prn = prn.strip()

            # Ignore completely empty rows
            if not name and not prn:
                continue

            if not valid_student(name, prn, class_name):
                invalid += 1
                continue

            try:
                conn.execute("""
                    INSERT INTO students (name, prn, class_name)
                    VALUES (?, ?, ?)
                """, (name, prn, class_name))

                added += 1

            except sqlite3.IntegrityError:
                duplicates += 1

        conn.commit()

    finally:
        conn.close()

    message = (
        f"Bulk registration completed! "
        f"Added: {added}, "
        f"Duplicates: {duplicates}, "
        f"Invalid rows: {invalid}."
    )

    return redirect(url_for("home", message=message))


# ==================================================
# BULK REGISTRATION - EXCEL UPLOAD
# ==================================================

@app.route("/upload-students", methods=["POST"])
def upload_students():

    class_name = request.form.get("excel_class", "").strip()
    file = request.files.get("student_file")

    if not class_name:
        return redirect(url_for(
            "home",
            message="Please enter the class name for Excel upload."
        ))

    if not file or not file.filename:
        return redirect(url_for(
            "home",
            message="Please select an Excel file."
        ))

    if not file.filename.lower().endswith(".xlsx"):
        return redirect(url_for(
            "home",
            message="Please upload an .xlsx Excel file."
        ))

    try:
        workbook = load_workbook(
            file,
            read_only=True,
            data_only=True
        )

        sheet = workbook.active

        # Read the first row as column headings
        headers = {}

        for index, cell in enumerate(sheet[1]):
            heading = str(cell.value or "").strip().lower()
            headers[heading] = index

        name_index = None
        prn_index = None

        for heading in ("student name", "name"):
            if heading in headers:
                name_index = headers[heading]
                break

        for heading in ("prn", "prn number", "roll number"):
            if heading in headers:
                prn_index = headers[heading]
                break

        if name_index is None or prn_index is None:
            workbook.close()

            return redirect(url_for(
                "home",
                message=(
                    "Excel headings must include Student Name "
                    "and PRN."
                )
            ))

        conn = get_db()

        added = 0
        duplicates = 0
        invalid = 0

        try:
            for row in sheet.iter_rows(min_row=2, values_only=True):

                name_value = (
                    row[name_index]
                    if name_index < len(row) else None
                )

                prn_value = (
                    row[prn_index]
                    if prn_index < len(row) else None
                )

                if name_value is None and prn_value is None:
                    continue

                name = str(name_value or "").strip()

                if isinstance(prn_value, float) and prn_value.is_integer():
                    prn = str(int(prn_value))
                elif prn_value is None:
                    prn = ""
                else:
                    prn = str(prn_value).strip()

                if not valid_student(name, prn, class_name):
                    invalid += 1
                    continue

                try:
                    conn.execute("""
                        INSERT INTO students
                        (name, prn, class_name)
                        VALUES (?, ?, ?)
                    """, (name, prn, class_name))

                    added += 1

                except sqlite3.IntegrityError:
                    duplicates += 1

            conn.commit()

        finally:
            conn.close()
            workbook.close()

        message = (
            f"Excel import completed! "
            f"Added: {added}, "
            f"Duplicates: {duplicates}, "
            f"Invalid rows: {invalid}."
        )

    except Exception as error:
        message = "Unable to read Excel file. Check its format."

    return redirect(url_for("home", message=message))


# ==================================================
# FIND STUDENT USING VOICE PRN
# ==================================================

@app.route("/api/student/<prn>")
def find_student(prn):

    conn = get_db()

    student = conn.execute("""
        SELECT id, name, prn, class_name
        FROM students
        WHERE prn = ?
    """, (prn,)).fetchone()

    conn.close()

    if student is None:
        return jsonify({
            "success": False,
            "message": "Student not found."
        }), 404

    return jsonify({
        "success": True,
        "student": dict(student)
    })


# ==================================================
# MARK ATTENDANCE
# ==================================================

@app.route("/api/attendance", methods=["POST"])
def mark_attendance():

    data = request.get_json(silent=True) or {}
    prn = str(data.get("prn", "")).strip()

    if not re.fullmatch(r"\d+", prn):
        return jsonify({
            "success": False,
            "message": "Invalid PRN."
        }), 400

    conn = get_db()

    student = conn.execute("""
        SELECT id FROM students WHERE prn = ?
    """, (prn,)).fetchone()

    if student is None:
        conn.close()

        return jsonify({
            "success": False,
            "message": "Student not found."
        }), 404

    now = datetime.now()

    try:
        conn.execute("""
            INSERT INTO attendance (
                student_id,
                attendance_date,
                attendance_time
            )
            VALUES (?, ?, ?)
        """, (
            student["id"],
            now.strftime("%Y-%m-%d"),
            now.strftime("%H:%M:%S")
        ))

        conn.commit()
        conn.close()

        return jsonify({
            "success": True,
            "message": "Attendance marked successfully!"
        })

    except sqlite3.IntegrityError:
        conn.close()

        return jsonify({
            "success": False,
            "message": "Attendance already marked today."
        }), 409


# ==================================================
# EXPORT ATTENDANCE TO EXCEL
# ==================================================

@app.route("/export-attendance")
def export_attendance():

    conn = get_db()

    records = conn.execute("""
        SELECT
            students.name,
            students.prn,
            students.class_name,
            attendance.attendance_date,
            attendance.attendance_time
        FROM attendance
        JOIN students
            ON students.id = attendance.student_id
        ORDER BY attendance.attendance_date DESC,
                 attendance.attendance_time DESC
    """).fetchall()

    conn.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Attendance Report"

    headers = [
        "Student Name",
        "PRN",
        "Class",
        "Attendance Date",
        "Attendance Time"
    ]

    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            start_color="198754",
            end_color="198754",
            fill_type="solid"
        )
        cell.alignment = Alignment(horizontal="center")

    for record in records:
        sheet.append([
            record["name"],
            record["prn"],
            record["class_name"],
            record["attendance_date"],
            record["attendance_time"]
        ])

    widths = [25, 18, 22, 20, 20]

    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[
            chr(64 + index)
        ].width = width

    # Keep PRNs as text in Excel
    for cell in sheet["B"][1:]:
        cell.number_format = "@"

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Attendance_Report.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ==================================================
# HTML DASHBOARD
# ==================================================

PAGE = """
<!DOCTYPE html>
<html>
<head>
    <title>Smart Voice Attendance System</title>

    <meta name="viewport" content="width=device-width, initial-scale=1">

    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f2f5f9;
            margin: 0;
            padding: 20px;
        }

        .container {
            max-width: 1100px;
            margin: auto;
        }

        h1 {
            text-align: center;
            color: #17365d;
        }

        .card {
            background: white;
            padding: 22px;
            margin: 20px 0;
            border-radius: 10px;
            box-shadow: 0 2px 8px #00000012;
        }

        input {
            display: block;
            width: 100%;
            box-sizing: border-box;
            padding: 11px;
            margin: 9px 0;
            border: 1px solid #ccc;
            border-radius: 5px;
        }

        button, .download-btn {
            display: inline-block;
            padding: 11px 17px;
            background: #1769aa;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            text-decoration: none;
            margin: 5px 3px;
        }

        button:hover, .download-btn:hover {
            opacity: 0.85;
        }

        .green {
            background: #198754;
        }

        .red {
            background: #dc3545;
        }

        .message {
            background: #e7f3ff;
            padding: 12px;
            border-radius: 5px;
            margin-top: 15px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
            background: white;
        }

        th, td {
            padding: 11px;
            border: 1px solid #ddd;
            text-align: left;
        }

        th {
            background: #1769aa;
            color: white;
        }

        .status {
            margin-top: 12px;
            font-weight: bold;
        }

        .student-row {
            display: grid;
            grid-template-columns: 2fr 1fr auto;
            gap: 10px;
            align-items: center;
        }

        .student-row input {
            min-width: 0;
        }

        .student-row button {
            margin-bottom: 9px;
        }

        .hint {
            color: #666;
            font-size: 14px;
        }

        @media(max-width: 650px) {
            body {
                padding: 8px;
            }

            .card {
                padding: 12px;
                overflow-x: auto;
            }

            table {
                min-width: 550px;
            }

            .student-row {
                grid-template-columns: 1fr 1fr auto;
                gap: 5px;
            }
        }
    </style>
</head>

<body>
<div class="container">

    <h1>Smart Voice-Based Student Attendance System</h1>

    {% if message %}
        <div class="message">{{ message }}</div>
    {% endif %}


    <!-- SINGLE STUDENT REGISTRATION -->

    <div class="card">
        <h2>Register One Student</h2>

        <form action="/register" method="POST">

            <input
                type="text"
                name="name"
                placeholder="Student Name"
                required
            >

            <input
                type="text"
                name="prn"
                placeholder="PRN Number"
                pattern="[0-9]+"
                required
            >

            <input
                type="text"
                name="class_name"
                placeholder="Class Name"
                required
            >

            <button type="submit" class="green">
                Register Student
            </button>

        </form>
    </div>


    <!-- BULK MANUAL REGISTRATION -->

    <div class="card">
        <h2>Bulk Student Registration</h2>

        <p class="hint">
            Enter the class once and add all students with their PRNs.
        </p>

        <form action="/bulk-register" method="POST">

            <label><b>Class Name</b></label>

            <input
                type="text"
                name="bulk_class"
                placeholder="Example: TY Computer"
                required
            >

            <h3>Student Details</h3>

            <div id="studentRows">

                <div class="student-row">
                    <input
                        type="text"
                        name="student_name"
                        placeholder="Student Name"
                    >

                    <input
                        type="text"
                        name="student_prn"
                        placeholder="PRN"
                        inputmode="numeric"
                    >

                    <button
                        type="button"
                        class="red"
                        onclick="removeRow(this)"
                    >
                        Remove
                    </button>
                </div>

                <div class="student-row">
                    <input
                        type="text"
                        name="student_name"
                        placeholder="Student Name"
                    >

                    <input
                        type="text"
                        name="student_prn"
                        placeholder="PRN"
                        inputmode="numeric"
                    >

                    <button
                        type="button"
                        class="red"
                        onclick="removeRow(this)"
                    >
                        Remove
                    </button>
                </div>

                <div class="student-row">
                    <input
                        type="text"
                        name="student_name"
                        placeholder="Student Name"
                    >

                    <input
                        type="text"
                        name="student_prn"
                        placeholder="PRN"
                        inputmode="numeric"
                    >

                    <button
                        type="button"
                        class="red"
                        onclick="removeRow(this)"
                    >
                        Remove
                    </button>
                </div>

            </div>

            <button
                type="button"
                onclick="addStudentRow()"
            >
                + Add Another Student
            </button>

            <br>

            <button type="submit" class="green">
                Register All Students
            </button>

        </form>
    </div>


    <!-- EXCEL STUDENT UPLOAD -->

    <div class="card">
        <h2>Import Students From Excel</h2>

        <p>
            Upload an Excel file containing all student names
            and PRNs for one class.
        </p>

        <p class="hint">
            Required Excel columns: Student Name and PRN.
        </p>

        <p class="hint">
            Example: Student Name | PRN
        </p>

        <form
            action="/upload-students"
            method="POST"
            enctype="multipart/form-data"
        >

            <label><b>Class Name</b></label>

            <input
                type="text"
                name="excel_class"
                placeholder="Example: TY Computer"
                required
            >

            <label><b>Select Excel File (.xlsx)</b></label>

            <input
                type="file"
                name="student_file"
                accept=".xlsx"
                required
            >

            <button type="submit" class="green">
                Upload and Register Students
            </button>

        </form>
    </div>


    <!-- REGISTERED STUDENTS -->

    <div class="card">
        <h2>All Registered Students</h2>

        <table>
            <tr>
                <th>Name</th>
                <th>PRN</th>
                <th>Class</th>
            </tr>

            {% for student in students %}
            <tr>
                <td>{{ student.name }}</td>
                <td>{{ student.prn }}</td>
                <td>{{ student.class_name }}</td>
            </tr>
            {% else %}
            <tr>
                <td colspan="3">No students registered yet.</td>
            </tr>
            {% endfor %}
        </table>

        <p>Total registered students: {{ students|length }}</p>
    </div>


    <!-- VOICE ATTENDANCE -->

    <div class="card">
        <h2>Voice-Based Attendance</h2>

        <p>
            Click Start Voice Recognition and speak your PRN
            digit by digit.
        </p>

        <button onclick="startRecognition()">
            Start Voice Recognition
        </button>

        <button onclick="stopRecognition()" class="red">
            Stop
        </button>

        <p id="recognizedText">
            Recognized PRN: Not yet recognized
        </p>

        <div id="studentDetails"></div>

        <button
            id="confirmButton"
            class="green"
            onclick="confirmAttendance()"
            style="display:none;"
        >
            Confirm Attendance
        </button>

        <p id="voiceStatus" class="status"></p>
    </div>


    <!-- TODAY'S ATTENDANCE -->

    <div class="card">
        <h2>Today's Attendance ({{ today }})</h2>

        <a href="/export-attendance" class="download-btn green">
            Download Attendance Excel
        </a>

        <table>
            <tr>
                <th>Name</th>
                <th>PRN</th>
                <th>Class</th>
                <th>Date</th>
                <th>Time</th>
            </tr>

            {% for record in attendance %}
            <tr>
                <td>{{ record.name }}</td>
                <td>{{ record.prn }}</td>
                <td>{{ record.class_name }}</td>
                <td>{{ record.attendance_date }}</td>
                <td>{{ record.attendance_time }}</td>
            </tr>
            {% else %}
            <tr>
                <td colspan="5">
                    No attendance marked today.
                </td>
            </tr>
            {% endfor %}
        </table>
    </div>

</div>


<script>

// ==================================================
// ADD OR REMOVE STUDENT ROWS
// ==================================================

function addStudentRow() {

    const container = document.getElementById("studentRows");

    const row = document.createElement("div");
    row.className = "student-row";

    const nameInput = document.createElement("input");
    nameInput.name = "student_name";
    nameInput.placeholder = "Student Name";

    const prnInput = document.createElement("input");
    prnInput.name = "student_prn";
    prnInput.placeholder = "PRN";
    prnInput.inputMode = "numeric";

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "red";
    removeButton.textContent = "Remove";
    removeButton.onclick = function() {
        removeRow(removeButton);
    };

    row.append(nameInput, prnInput, removeButton);
    container.appendChild(row);
}


function removeRow(button) {

    const container = document.getElementById("studentRows");

    if (container.children.length > 1) {
        button.parentElement.remove();
    } else {
        alert("Keep at least one student row.");
    }
}


// ==================================================
// VOICE RECOGNITION
// ==================================================

let recognition = null;
let currentPRN = "";

const digitWords = {
    "zero": "0",
    "oh": "0",
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9"
};


function convertToPRN(transcript) {

    let words = transcript.toLowerCase()
        .replace(/-/g, " ")
        .replace(/[^a-z0-9 ]/g, " ")
        .trim()
        .split(/\s+/);

    let digits = "";

    for (let word of words) {

        if (/^[0-9]$/.test(word)) {
            digits += word;

        } else if (digitWords[word] !== undefined) {
            digits += digitWords[word];

        } else {
            return "";
        }
    }

    return digits;
}


function startRecognition() {

    const SpeechRecognition =
        window.SpeechRecognition ||
        window.webkitSpeechRecognition;

    if (!SpeechRecognition) {
        document.getElementById("voiceStatus").textContent =
            "Speech recognition is not supported. Use Chrome.";

        return;
    }

    if (recognition) {
        recognition.stop();
    }

    recognition = new SpeechRecognition();

    recognition.lang = "en-IN";
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.maxAlternatives = 3;

    currentPRN = "";

    document.getElementById("confirmButton").style.display =
        "none";

    document.getElementById("studentDetails").textContent = "";

    document.getElementById("voiceStatus").textContent =
        "Listening... Please speak your PRN.";

    recognition.onresult = async function(event) {

        const transcript =
            event.results[0][0].transcript;

        const prn = convertToPRN(transcript);

        document.getElementById("recognizedText").textContent =
            "Recognized speech: " + transcript +
            " | PRN: " + (prn || "Could not recognize digits");

        if (!prn) {
            document.getElementById("voiceStatus").textContent =
                "Please speak the PRN digit by digit and try again.";

            return;
        }

        currentPRN = prn;

        await findStudent(prn);
    };


    recognition.onerror = function(event) {

        document.getElementById("voiceStatus").textContent =
            "Voice recognition error: " + event.error;
    };


    recognition.onend = function() {

        if (
            document.getElementById("voiceStatus").textContent ===
            "Listening... Please speak your PRN."
        ) {
            document.getElementById("voiceStatus").textContent =
                "Listening stopped. Try again.";
        }
    };

    recognition.start();
}


function stopRecognition() {

    if (recognition) {
        recognition.stop();
    }

    document.getElementById("voiceStatus").textContent =
        "Voice recognition stopped.";
}


// ==================================================
// FIND STUDENT
// ==================================================

async function findStudent(prn) {

    try {

        const response = await fetch(
            "/api/student/" + encodeURIComponent(prn)
        );

        const data = await response.json();

        if (!response.ok || !data.success) {

            document.getElementById("studentDetails").textContent =
                "Student not found for PRN: " + prn;

            document.getElementById("confirmButton").style.display =
                "none";

            document.getElementById("voiceStatus").textContent =
                "Please check the PRN and try again.";

            return;
        }

        const student = data.student;

        const details = document.getElementById("studentDetails");
        details.replaceChildren();

        const heading = document.createElement("h3");
        heading.textContent = "Student Identified";

        const name = document.createElement("p");
        name.textContent = "Name: " + student.name;

        const prnText = document.createElement("p");
        prnText.textContent = "PRN: " + student.prn;

        const classText = document.createElement("p");
        classText.textContent = "Class: " + student.class_name;

        details.append(heading, name, prnText, classText);

        document.getElementById("confirmButton").style.display =
            "inline-block";

        document.getElementById("voiceStatus").textContent =
            "Check student details and confirm attendance.";

    } catch (error) {

        document.getElementById("voiceStatus").textContent =
            "Unable to connect to the server.";
    }
}


// ==================================================
// CONFIRM ATTENDANCE
// ==================================================

async function confirmAttendance() {

    if (!currentPRN) {
        return;
    }

    try {

        const response = await fetch("/api/attendance", {

            method: "POST",

            headers: {
                "Content-Type": "application/json"
            },

            body: JSON.stringify({
                prn: currentPRN
            })
        });

        const data = await response.json();

        document.getElementById("voiceStatus").textContent =
            data.message;

        document.getElementById("confirmButton").style.display =
            "none";

        if (response.ok && data.success) {

            currentPRN = "";

            setTimeout(function() {
                window.location.reload();
            }, 1000);
        }

    } catch (error) {

        document.getElementById("voiceStatus").textContent =
            "Unable to save attendance. Please try again.";
    }
}

</script>

</body>
</html>
"""


# ==================================================
# START APPLICATION
# ==================================================

if __name__ == "__main__":
    init_db()
    app.run(debug=True)
