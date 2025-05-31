# app.py - Vercel entry point
import os
import sys

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import app

# Expose the Flask app for Vercel
app = app

if __name__ == "__main__":
    app.run(debug=True) 