
import os

CLIENT_ID = os.environ.get("MS_GRAPH_CLIENT_ID")
CLIENT_SECRET = os.environ.get("MS_GRAPH_CLIENT_SECRET")
TENANT_ID = os.environ.get("MS_GRAPH_TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "http://localhost:5000/callback"
SCOPE = ["Files.ReadWrite.All", "User.Read"]
