#!/usr/bin/env python3
"""
Simple HTTP server to serve the IC++ report with preset loading

Usage:
    python3 serve.py

Then open: http://localhost:8000/icpp_breakdown_report.html
"""

import http.server
import socketserver
import os

PORT = 8000

class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Add CORS headers to allow file loading
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET')
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

os.chdir('/Users/manik.soin/Desktop/merchant funding')

with socketserver.TCPServer(("", PORT), MyHTTPRequestHandler) as httpd:
    print(f"✓ Server running at http://localhost:{PORT}/")
    print(f"\n📊 Open in browser: http://localhost:{PORT}/icpp_breakdown_report.html")
    print(f"\n💡 Benefits of using server:")
    print(f"   - Preset loading works (dropdown + load button)")
    print(f"   - Drag & drop still works")
    print(f"   - All features fully functional")
    print(f"\nPress Ctrl+C to stop server")
    httpd.serve_forever()
