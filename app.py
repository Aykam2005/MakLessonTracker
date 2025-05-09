import os
import openpyxl
import msal
import requests
from flask import Flask, redirect, render_template, request, session, url_for
from config import CLIENT_ID, CLIENT_SECRET, AUTHORITY, REDIRECT_URI, SCOPE

app = Flask(__name__)
app.secret_key = os.urandom(24)


def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=authority or AUTHORITY,
        client_credential=CLIENT_SECRET, token_cache=cache)


def _build_auth_code_flow(scopes=None, redirect_uri=None):
    return _build_msal_app().initiate_auth_code_flow(
        scopes or [],
        redirect_uri=redirect_uri or REDIRECT_URI)


@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return redirect(url_for("lessons"))


@app.route("/login")
def login():
    session["flow"] = _build_auth_code_flow(scopes=SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(session["flow"]["auth_uri"])


@app.route("/callback")
def authorized():
    try:
        cache = msal.SerializableTokenCache()
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args)
    except ValueError:
        return "Authentication failed.", 400

    if "error" in result:
        return f"Error: {result['error']} - {result.get('error_description')}", 400

    session["user"] = result.get("id_token_claims")
    session["access_token"] = result["access_token"]
    return redirect(url_for("lessons"))


@app.route("/lessons")
def lessons():
    if "access_token" not in session:
        return redirect(url_for("login"))

    token = session["access_token"]
    headers = {"Authorization": f"Bearer {token}"}
    file_url = "https://graph.microsoft.com/v1.0/me/drive/root:/LessonTracker.xlsx:/content"
    response = requests.get(file_url, headers=headers)

    if response.status_code != 200:
        return f"Failed to download Excel file: {response.text}", 400

    with open("LessonTracker.xlsx", "wb") as f:
        f.write(response.content)

    workbook = openpyxl.load_workbook("LessonTracker.xlsx")
    sheet = workbook["lesson log"]
    lessons = [
        [cell.value for cell in row]
        for row in sheet.iter_rows(min_row=2, values_only=True)
    ]
    return render_template("lessons.html", lessons=lessons)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('index', _external=True)}"
    )


if __name__ == "__main__":
    app.run(debug=True)
