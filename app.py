from flask import Flask, session, redirect, url_for, request, render_template
import msal
import requests
from openpyxl import load_workbook
import os
from config import CLIENT_ID, CLIENT_SECRET, TENANT_ID, AUTHORITY, REDIRECT_URI, SCOPE

app = Flask(__name__)
app.secret_key = os.urandom(24)

def _build_msal_app():
    """Build MSAL confidential client app."""
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

def _build_auth_code_flow():
    """Initiate the auth code flow."""
    return _build_msal_app().initiate_auth_code_flow(
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )

@app.route("/")
def index():
    """Main page after login."""
    if "token" not in session:
        return redirect(url_for("login"))
    return render_template("index.html")

@app.route("/login")
def login():
    """Login route: redirects user to Microsoft login page."""
    session["flow"] = _build_auth_code_flow()
    return redirect(session["flow"]["auth_uri"])

@app.route("/callback")
def callback():
    """Callback URL after authentication with MS."""
    flow = session.get("flow", {})
    
    # Check if authorization code is provided by Microsoft
    if "code" not in request.args:
        return "Missing 'code' parameter in callback", 400
    
    # Get the access token using the authorization code
    result = _build_msal_app().acquire_token_by_authorization_code(
        request.args["code"],
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
    
    if "access_token" in result:
        session["token"] = result["access_token"]
        return redirect(url_for("index"))
    
    return f"Login failed: {result.get('error_description')}", 500

@app.route("/read_excel")
def read_excel():
    """Route to read Excel file from OneDrive."""
    if "token" not in session:
        return redirect(url_for("login"))
    
    file_path = get_excel_file_from_onedrive(session["token"], "LessonTracker.xlsm")
    
    if file_path:
        try:
            wb = load_workbook(file_path, keep_vba=True)
            sheet = wb["lesson log"]
            last_row = sheet.max_row
            last_entry = [cell.value for cell in sheet[last_row]]
            return f"Last lesson entry: {last_entry}"
        except Exception as e:
            return f"Failed to read Excel file: {str(e)}", 500
    return "Failed to load Excel file.", 500

def get_excel_file_from_onedrive(access_token, file_name):
    """Get Excel file from OneDrive."""
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f'https://graph.microsoft.com/v1.0/me/drive/root:/{file_name}:/content'

    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        local_file = f"temp_{file_name}"
        with open(local_file, 'wb') as f:
            f.write(response.content)
        return local_file
    else:
        print(f"Graph API error: {response.status_code} - {response.text}")
        return None

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5001, debug=True)
