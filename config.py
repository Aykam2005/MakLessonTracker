
import os

CLIENT_ID = os.environ.get("MS_GRAPH_CLIENT_ID")
CLIENT_SECRET = os.environ.get("MS_GRAPH_CLIENT_SECRET")
TENANT_ID = "your-actual-tenant-id"  # <- this should be a valid GUID like '8f46243b-217e-4fe7-bd5c-a97939a6d86d'
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = "http://localhost:5000/callback"
SCOPE = ["Files.ReadWrite.All", "User.Read"]
