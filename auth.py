# auth.py — Browser opens automatically

import msal
import json

CLIENT_ID = "9e8a0a43-04f4-4019-afe3-a3d837bb9e4d"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Mail.ReadWrite", "Mail.ReadWrite.Shared", "Mail.Send"]
TOKEN_CACHE_FILE = "token_cache.json"

def authenticate():
    cache = msal.SerializableTokenCache()

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    # This automatically opens your browser to login!
    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" in result:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())
        print("✅ Login successful! Token saved.")
    else:
        print("❌ Error:", result.get("error_description"))

if __name__ == "__main__":
    authenticate()
