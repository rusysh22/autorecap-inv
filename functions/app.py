import os
import sys

# Add root directory to path so we can import app.py
root_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.append(root_path)

from app import app
import serverless_wsgi

def handler(event, context):
    return serverless_wsgi.handle_request(app, event, context)
