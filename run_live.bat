@echo off
echo Starting Instagram IST Tracker & Cloudflare Live Tunnel...
start /B python -m streamlit run instagram_ist_tracker_app.py --server.headless true --server.port 8501 --server.address 0.0.0.0
timeout /t 3 /nobreak >nul
echo Tunneling to public HTTPS...
cloudflared.exe tunnel --url http://127.0.0.1:8501
pause
