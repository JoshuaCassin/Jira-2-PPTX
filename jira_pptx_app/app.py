from flask import Flask, render_template, request, send_from_directory, jsonify
from jira_api import fetch_jira_data
from csv_utils import generate_csv_from_issues 
from pptx_utils import modify_pptx
import os
from werkzeug.utils import secure_filename
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route("/", methods=['GET'])
def index():
    return render_template('index.html')

@app.route("/generate_pptx", methods=['POST'])
def generate_pptx_route():
    try:
        # Retrieve form data
        email = request.form.get('email')
        api_key = request.form.get('api_key')
        jql_query = request.form.get('jql_query')
        fields_to_include = [field.strip() for field in request.form.get('fields_to_include').split(',')]
        fix_version_ce = request.form.get('fix_version_ce')
        fix_version_ee = request.form.get('fix_version_ee')
        jira_server_url = request.form.get('jira_server_url')

        # Handling the PowerPoint file upload
        pptx_template = request.files.get('pptx_template')
        if pptx_template:
            filename = secure_filename(pptx_template.filename)  # Ensure the filename is safe to use
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            pptx_template.save(save_path)
        else:
            return "PowerPoint file upload failed", 400

        response = fetch_jira_data(jira_server_url, email, api_key, jql_query)
        print("Response type:", type(response))

        # Check if there's an error in the fetched data
        if isinstance(response, dict) and 'error' in response:
            print(response['error'])
            return jsonify({"error": response['error']}), 500

        if isinstance(response, list):
            print(response[:5])  # print the first 5 items
        elif isinstance(response, dict):
            print({k: response[k] for k in list(response)[:6]})  # print the first 6 key-value pairs
        else:
            print("Response:", response)

        # If no errors, continue processing
        # Generate the CSV content as a string
        csv_content = generate_csv_from_issues(response, fields_to_include, fix_version_ce, fix_version_ee)

        # Create a save path for the modified presentation
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        modified_filename = f"modified_{timestamp}.pptx"
        modified_save_path = os.path.join(app.config['UPLOAD_FOLDER'], modified_filename)

        # Modify the PPTX based on this CSV
        rows_per_slide = int(request.form.get('rows_per_slide', 12))
        modify_pptx(save_path, fix_version_ce, fix_version_ee, fields_to_include, rows_per_slide, output_path=modified_save_path)

        # Delete the uploaded PowerPoint file
        os.remove(save_path)


        # Return the display_response template
        return render_template('display_response.html', csv_data=csv_content, download_url=f"/download/{modified_filename}")

    except RuntimeError as e:
        app.logger.error(str(e))
        return "Error fetching data from Jira", 500




@app.route("/create_pptx", methods=['POST'])
def create_pptx_route():
    # 1. Handling the PowerPoint file upload
    pptx_template = request.files.get('pptx_template')
    if not pptx_template:
        return "PowerPoint file upload failed", 400
    
    filename = secure_filename(pptx_template.filename)
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    pptx_template.save(save_path)

    # 2. Create a save path for the modified presentation
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    modified_filename = f"modified_{timestamp}.pptx"
    modified_save_path = os.path.join(app.config['UPLOAD_FOLDER'], modified_filename)

    # 3. Generate the CSV again (Note: You need to get the response, fields_to_include, fix_version_ce, fix_version_ee from somewhere. Consider sending these as hidden fields in your HTML form).
    response = ... # fetch or get it from the form or somewhere
    fields_to_include = ... # fetch or get it from the form
    fix_version_ce = ... # fetch or get it from the form
    fix_version_ee = ... # fetch or get it from the form
    generate_csv_from_response(response, fields_to_include, fix_version_ce, fix_version_ee)

    # 4. Modify the PPTX based on this CSV
    rows_per_slide = 10  # Adjust this value as per your requirement
    modify_pptx(save_path, fix_version_ce, fix_version_ee, fields_to_include, rows_per_slide, output_path=modified_save_path)

    # 5. Delete the uploaded PowerPoint file
    os.remove(save_path)

    # 6. Return a download link for the modified PPTX file
    return f"Your PPTX file is ready! <a href='/download/{modified_filename}'>Click here to download.</a>"



@app.route('/download/<string:pptx_template_filename>')
def download_file(pptx_template_filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], pptx_template_filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=56000)
