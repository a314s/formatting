@echo off
echo Starting Excel Formatter Server...
echo.
echo The server will be available at http://localhost:8001
echo.
echo To stop the server:
echo 1. Use the "Turn Off Server" button in the web interface, or
echo 2. Press Ctrl+C in this window and confirm with Y
echo.
python server.py
echo.
echo Server has been shut down.
pause