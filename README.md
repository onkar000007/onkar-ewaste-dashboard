# Onkar Shelke — Assignment Dashboard

> **Student:** Onkar Shelke (Roll No: `25102C2002`)  
> **Course:** E-Waste & Environmental Management (EWEM) · Semester 5  
> **Institution:** Vidyalankar Institute of Technology (VIT), Mumbai

An interactive academic dashboard featuring live in-browser document previews (DOCX, PPTX, XLSX), native video playback, and academic review rubrics for all 15 course assignments.

---

## 🚀 Hosting on Vercel

This repository is pre-configured for **Zero-Configuration Static Hosting on Vercel**.

### Method 1: Deploy via GitHub (Recommended)
1. Initialize Git and commit the files (if not already done):
   ```bash
   git init
   git add .
   git commit -m "Initial commit for Vercel deployment"
   ```
2. Push your repository to GitHub / GitLab / Bitbucket.
3. Go to [vercel.com/new](https://vercel.com/new).
4. Click **Import** on your repository.
5. In **Framework Preset**, select **Other** (Root directory: `./`).
6. Click **Deploy**. Your dashboard will be live at `https://your-project.vercel.app` in under 30 seconds!

---

### Method 2: Deploy via Vercel CLI
1. Open PowerShell or Command Prompt in this folder.
2. Run:
   ```bash
   npx vercel
   ```
3. Follow the quick prompts in your terminal:
   - Set up and deploy? **Y**
   - Which scope? (Select your Vercel account)
   - Link to existing project? **N**
   - Project name? **onkar-shelke-assignment-dashboard**
   - Directory is `./`? **Y**
   - Want to modify settings? **N**
4. For production deployment, run:
   ```bash
   npx vercel --prod
   ```

---

## 📁 Project & File Structure

```text
├── index.html               # Main dashboard entry point
├── 404.html                 # Custom dark glassmorphic 404 page
├── favicon.svg              # Dashboard favicon icon
├── vercel.json              # Vercel configuration (caching, security, video headers)
├── .vercelignore            # Prevents redundant file uploads to keep deployments fast
├── .gitignore               # Git ignore rules (.vercel/, node_modules/, etc.)
├── package.json             # Project metadata & local dev scripts
├── assets/
│   ├── assignment-files/    # 27 source assignment documents (DOCX, PPTX, XLSX, PDF, MP4)
│   └── previews/            # HTML document previews with extracted figures
│       └── media/           # Extracted document images & diagrams
├── assignments/             # Local assignment folders (kept intact)
├── files/                   # Local subject category folders (kept intact)
└── build_previews.py        # Python preview generator script
```

---

## 💻 Local Testing & Development

To test the website locally before deploying:

```bash
# Option 1: Using Node / npx
npx serve

# Option 2: Using Python built-in server
python -m http.server 3000
```

Open `http://localhost:3000` in your web browser.
