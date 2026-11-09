# Instagram Post & Reel Timestamp Tracker (IST)

An interactive Streamlit web dashboard and Excel exporter that converts Instagram post and Reel timestamps to **Indian Standard Time (IST / UTC+5:30)**.

---

## 🚀 Instant Deployment Options

### Option 1: Streamlit Community Cloud (Recommended — Free & 24/7 Live)
1. Push this repository to your GitHub account:
   ```bash
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo-name>.git
   git push -u origin main
   ```
2. Go to **[share.streamlit.io](https://share.streamlit.io)** and log in with GitHub.
3. Click **"New app"**:
   - Repository: `<your-username>/<your-repo-name>`
   - Branch: `main`
   - Main file path: `instagram_ist_tracker_app.py`
4. Click **"Deploy"**. Streamlit Cloud will give you a public URL (e.g., `https://your-app.streamlit.app`) you can open from Chrome on any device (phone, laptop, iPad).

---

### Option 2: Instant Phone/Device Access on Same Wi-Fi (No Cloud Setup Needed)
If your other device (e.g. phone) is connected to the same Wi-Fi router as this computer:
1. Run the app:
   ```bash
   streamlit run instagram_ist_tracker_app.py --server.address 0.0.0.0
   ```
2. Streamlit will display a **Network URL** (e.g. `http://192.168.1.X:8501`).
3. Open that URL directly in Chrome on your phone or other device.

---

### Option 3: Instant Public Web Link with Cloudflare Tunnel
To share or access from anywhere in the world while running locally:
```bash
# If using cloudflared:
cloudflared tunnel --url http://localhost:8501
```
Or with ngrok:
```bash
ngrok http 8501
```

---

> **Note on Instagram Scraping:**
> Instagram occasionally rate-limits or challenges public cloud IPs (AWS/GCP/Streamlit Cloud). If an extraction stops early or returns an error on cloud hosting, simply enter a secondary/burner Instagram account under **Optional Instagram Login** in the sidebar.
