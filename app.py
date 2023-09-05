from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os



app = Flask(__name__)

# Define a function to read data from Excel
def read_excel_data():
    try:
        if os.path.exists('course_requests.xlsx'):
            df = pd.read_excel('course_requests.xlsx', sheet_name='Course Requests')
            return df.to_dict(orient='records')
        else:
            return []
    except Exception as e:
        print(f"Error reading Excel data: {str(e)}")
        return []

@app.route('/')
def dashboard():
    # Replace with your logic to fetch and display course cards
    return render_template('dashboard.html')

@app.route('/request_training', methods=['GET', 'POST'])
def request_training():
    if request.method == 'POST':
        # Get form data
        course_type = request.form.get('course_type')
        course_name = request.form.get('course_name')
        course_duration = request.form.get('course_duration')
        online_offline = request.form.get('online_offline')
        team_selection = request.form.get('team_selection')

        # Store data in a new Excel file
        data = {
            'Course Type': [course_type],
            'Course Name': [course_name],
            'Course Duration': [course_duration],
            'Online/Offline': [online_offline],
            'Team Selection': [team_selection]
        }
        df = pd.DataFrame(data)

        # Write to a new Excel file
        # with pd.ExcelWriter('course_requests.xlsx', engine='ExcelWriter', mode='w') as writer:
        with pd.ExcelWriter('course_requests.xlsx',mode='a',if_sheet_exists="overlay") as writer:
            df.to_excel(writer, sheet_name="Course Requests",header=None, startrow=writer.sheets["Course Requests"].max_row,index=False)

        # Redirect to the notification page
        return redirect(url_for('notification'))

    return render_template('request_form.html')

@app.route('/notification')
def notification():
    return render_template('notification.html')

@app.route('/requested_courses')
def requested_courses():
    # Fetch data from Excel and pass it to the template
    requested_courses = read_excel_data()
    return render_template('requested_courses.html', requested_courses=requested_courses)

if __name__ == '__main__':
    app.run(debug=True)