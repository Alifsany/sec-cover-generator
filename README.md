# Assignment Cover Page Generator — Shyamoli Engineering College

A mobile-friendly, production-ready web application for generating, previewing, and downloading official Shyamoli Engineering College (affiliated with University of Dhaka) assignment cover pages in both **editable Word (`.docx`)** and **print-ready single-page PDF** formats.

---

## 🚀 Features

- **Mobile First Design**: Touch-friendly inputs, adaptive responsive layout, zero horizontal scroll across mobile viewports (320px–414px), tablets, and desktop computers.
- **Genuine Editable DOCX**: Generates real Microsoft Word documents (`.docx`) with 14 pt typography for student/teacher details, centered header watermark crest, and table layout.
- **Single-Page PDF**: Vector-rendered single-page PDF generated via ReportLab matching the Word layout and original geometry.
- **Fixed College Branding**: College name `SHYAMOLI ENGINEERING COLLEGE` is strictly fixed; student fields are blank by default with helpful placeholders.
- **No Personal Information**: Clean, generic templates and forms without prefilled personal information.
- **Privacy First**: No account registration, no database storage, and no third-party advertisements.

---

## 🛠️ Local Development & Running

### Requirements
- Python 3.9+ installed
- Dependencies in `requirements.txt`

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Start the local server
```bash
python app.py
```

### 3. Open in Browser
Visit:
```
http://localhost:5000
```
or from any phone on the same Wi-Fi using `http://<your-computer-ip>:5000`.

---

## 🌐 Step-by-Step Render Deployment Guide

Follow these steps to publish the application online for free via **GitHub** and **Render**.

### Step 1: Create a GitHub Repository
1. Log in to [GitHub](https://github.com/).
2. In the top right corner, click **+** and select **New repository**.
3. Choose a Repository name (e.g. `sec-cover-generator`).
4. Set visibility to **Public** (or **Private**).
5. **Do not** check "Add a README file" or ".gitignore" (these are already configured in the project).
6. Click **Create repository**.
7. Copy the repository URL (e.g., `https://github.com/<your-username>/sec-cover-generator.git`).

### Step 2: Push the Code from Windows PowerShell
Open PowerShell in `C:\cover page` and run:
```powershell
cd "C:\cover page"
git add .
git commit -m "Production release for Render deployment"
git branch -M main
git remote add origin https://github.com/<your-username>/sec-cover-generator.git
git push -u origin main
```
*(Replace `<your-username>` and `sec-cover-generator` with your actual GitHub username and repository name).*

### Step 3: Create and Configure a Web Service on Render
1. Log in to your [Render Dashboard](https://dashboard.render.com/).
2. Click **New +** (top right) and select **Web Service**.
3. Choose **Build and deploy from a Git repository** and connect your GitHub account.
4. Locate and select your newly pushed repository (`sec-cover-generator`).
5. Configure the service settings:
   - **Name**: `sec-cover-generator` (or your chosen service name)
   - **Region**: Select the region nearest to your users (e.g., *Singapore*, *Frankfurt*, or *Oregon*)
   - **Branch**: `main`
   - **Root Directory**: *(Leave blank)*
   - **Runtime**: `Python`
   - **Build Command**:
     ```bash
     pip install -r requirements.txt
     ```
   - **Start Command**:
     ```bash
     gunicorn app:app
     ```
   - **Instance Type**: `Free`
6. Click **Deploy Web Service** at the bottom.

### Step 4: Monitor Build Logs & Verify Deployment
1. Render will start the build process automatically.
2. The logs will display:
   - Installing packages from `requirements.txt` (`Flask`, `python-docx`, `ReportLab`, `Pillow`, `gunicorn`).
   - Starting Gunicorn worker processes.
   - `Listening at: http://0.0.0.0:10000` (Render dynamically sets the `PORT`).
3. Once the status shows **Live**, click the URL provided at the top (e.g. `https://sec-cover-generator.onrender.com`).
4. Open the website on your smartphone or computer, fill in the fields, and verify both PDF and DOCX downloads.

---

## 🔍 Troubleshooting & Production FAQ

- **Port Binding Issues**: The application dynamically binds to `0.0.0.0` using `os.environ.get('PORT', 5000)`. Render passes its allocated port automatically.
- **Missing Asset Errors**: `doc_generator.py` uses `find_asset()` to resolve fonts and watermark graphics across root and `static/` folders regardless of the current working directory.
- **DOCX / PDF Download on Mobile Safari**: If iOS Safari suppresses direct blob downloads, a persistent tap link automatically appears below the action buttons allowing students to view or save the file to Apple Files or Google Drive.
- **Cold Starts on Free Tier**: On Render's free tier, the instance will spin down after 15 minutes of inactivity. When a student visits the link after inactivity, the first load may take ~30–50 seconds to boot up. Subsequent interactions are instant.
