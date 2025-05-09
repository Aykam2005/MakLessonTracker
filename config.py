# config.py

CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret"
AUTHORITY = "https://login.microsoftonline.com/{your-tenant-id}"  # Replace with your tenant ID
REDIRECT_URI = "http://localhost:5000/callback"
SCOPE = ["User.Read", "Files.Read"]
