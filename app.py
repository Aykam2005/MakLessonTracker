import os
import openpyxl
import msal
import requests
from flask import Flask, redirect, render_template, request, session, url_for
from io import BytesIO  # Required to read Excel file from memory

from config import CLIENT_ID, CLIENT_SECRET, AUTHORITY, REDIRECT_URI, SCOPE

app = Flask(__name__)
app.secret_key = os.urandom(24)


def _build_msal_app(cache=None, authority=None):
    """Builds and returns the MSAL ConfidentialClientApplication."""
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority or AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )


def _build_auth_code_flow(scopes=None, redirect_uri=None):
    """Initiates the auth code flow for the given scopes and redirect URI."""
    return _build_msal_app().initiate_auth_code_flow(
        scopes or [], redirect_uri=redirect_uri or REDIRECT_URI
    )


@app.route("/")
def index():
    """Main route: checks if user is authenticated."""
    if not session.get("user"):
        return redirect(url_for("login"))
    return redirect(url_for("lessons"))


@app.route("/login")
def login():
    """Initiates the login process by redirecting to the Microsoft login page."""
    session["flow"] = _build_auth_code_flow(scopes=SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(session["flow"]["auth_uri"])


@app.route("/callback")
def authorized():
    """Handles the callback after successful authentication."""
    cache = msal.SerializableTokenCache()

    try:
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args
        )
    except ValueError:
        return "Authentication failed. Invalid response.", 400

    if "error" in result:
        return f"Error: {result['error']} - {result.get('error_description')}", 400

    # Save user and access token in session
    session["user"] = result.get("id_token_claims")
    session["access_token"] = result.get("access_token")
    return redirect(url_for("lessons"))


@app.route("/lessons")
def lessons():
    """Fetches and displays lessons from the Excel file stored in OneDrive."""
    if "access_token" not in session:
        return redirect(url_for("login"))

    token = session["access_token"]
    headers = {"Authorization": f"Bearer {token}"}
    file_url = "https://graph.microsoft.com/v1.0/me/drive/root:/LessonTracker.xlsx:/content"

    response = requests.get(file_url, headers=headers)

    if response.status_code != 200:
        return f"Failed to download Excel file: {response.text}", 400

    # Load workbook from response content
    try:
        workbook = openpyxl.load_workbook(filename=BytesIO(response.content))
        sheet = workbook["lesson log"]
    except Exception as e:
        return f"Error reading Excel file: {e}", 500

    # Extract lessons starting from row 2 (skipping headers)
    lessons = [
        list(row)
        for row in sheet.iter_rows(min_row=2, values_only=True)
        if any(row)  # skip completely empty rows
    ]

    return render_template("lessons.html", lessons=lessons)


@app.route("/logout")
def logout():
    """Logs the user out by clearing the session and redirecting to Microsoft logout."""
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('index', _external=True)}"
    )


if __name__ == "__main__":
    app.run(debug=True)
