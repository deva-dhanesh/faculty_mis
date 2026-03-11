"""
serve.py – Start the Faculty MIS with Waitress (Windows-compatible WSGI server)

Usage:
    & .\\venv\\Scripts\\python.exe serve.py

The app will be reachable at:
    Local  : http://127.0.0.1:5000
    Network: http://<your-ip>:5000
"""

from waitress import serve
from app import app
import os
from dotenv import load_dotenv

load_dotenv()

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", 5000))
THREADS = int(os.environ.get("THREADS", 4))

if __name__ == "__main__":
    print(f"[waitress] Serving on http://{HOST}:{PORT}  (threads={THREADS})")
    print("[waitress] Press Ctrl+C to stop.\n")
    serve(app, host=HOST, port=PORT, threads=THREADS)
