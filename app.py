import pandas as pd
from flask import Flask, send_file, render_template
import os
from io import BytesIO

app = Flask(__name__)

# Configure upload folder
app.config['UPLOAD_FOLDER'] = 'static'
os.makedirs(os.path.join(app.config['UPLOAD_FOLDER'], 'reports'), exist_ok=True)
os.makedirs(os.path.join(app.config['UPLOAD_FOLDER'], 'employees'), exist_ok=True)

def generate_data():
    """Generate sample data and Excel files"""
    # Sample data - same as your original
    pattama_data = {
        "Date": ["15-Jan-2025", "15-Jan-2025"],
        "Candidate Name": ["Mr.A", "Mr.B"],
        "Role": ["Data Analyst", "Web Developer"],
        "Interview": ["Yes", "Yes"],
        "Status": ["Pass", "Fail"],
        "Remark": ["", ""]
    }
    
    raewwadee_data = {
        "Date": ["15-Jan-2025", "15-Jan-2025"],
        "Candidate Name": ["Mr.C", "Mr.D"],
        "Role": ["Software Tester", "Project Coordinator"],
        "Interview": ["Yes", "Yes"],
        "Status": ["Fail", "Pass"],
        "Remark": ["", ""]
    }
    
    new_employee_data = {
        "Employee Name": ["Mr.A", "Mr.D"],
        "Join Date": ["3-Feb-2025", "17-Feb-2025"],
        "Role": ["Data Analyst", "Project Coordinator"],
        "DOB": ["01-01-2000", "01-12-2001"],
        "ID Card": ["1-1111-11111-11-1", "2-2222-2222-22-2"],
        "Remark": ["", ""]
    }

    # Save to files
    pd.DataFrame(pattama_data).to_excel('static/reports/Daily_report_Pattama.xlsx', index=False)
    pd.DataFrame(raewwadee_data).to_excel('static/reports/Daily_report_Raewwadee.xlsx', index=False)
    pd.DataFrame(new_employee_data).to_excel('static/employees/New_Employees.xlsx', index=False)

    return create_dashboard()

def create_dashboard():
    """Process data and create final dashboard"""
    pattama = pd.read_excel('static/reports/Daily_report_Pattama.xlsx')
    raewwadee = pd.read_excel('static/reports/Daily_report_Raewwadee.xlsx')
    new_employees = pd.read_excel('static/employees/New_Employees.xlsx')
    
    pattama["Interviewer"] = "Pattama Sooksan"
    raewwadee["Interviewer"] = "Raewwadee Jaidee"
    all_interviews = pd.concat([pattama, raewwadee])
    
    result = pd.merge(
        new_employees,
        all_interviews[["Candidate Name", "Role", "Interviewer"]],
        left_on=["Employee Name", "Role"],
        right_on=["Candidate Name", "Role"],
        how="left"
    )
    final = result[["Employee Name", "Join Date", "Role", "Interviewer"]].dropna()
    final.to_excel('static/employees/Employee_Dashboard.xlsx', index=False)
    return final

@app.route("/")
def dashboard():
    df = generate_data()
    return render_template('dashboard.html', 
                         tables=[df.to_html(classes='data')],
                         titles=df.columns.values)

@app.route("/download")
def download():
    return send_file('static/employees/Employee_Dashboard.xlsx',
                     as_attachment=True,
                     download_name='Employee_Dashboard.xlsx')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))