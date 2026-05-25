# Academic Hub

A dashboard to track scholarships, university applications, and professor emails — all in one place.

---

## What You Need Before Starting

This app runs on your computer using **Node.js**. Think of Node.js like a engine that powers the app.

1. **Download Node.js**
   - Go to https://nodejs.org
   - Download the **LTS** version (the big button on the left)
   - Open the downloaded file and follow the installation steps (just keep clicking "Next")
   - ✅ Done — this also installs **npm** (the package manager) automatically

2. **Download this project**
   - Click the green "Code" button on this repo's GitHub page
   - Click **"Download ZIP"**
   - Extract (unzip) the folder somewhere you'll remember, like your Desktop

---

## How to Run the App

You'll use a program called **Terminal** (Mac) or **Command Prompt** (Windows) to type a few commands.

### Open Terminal / Command Prompt

- **Mac**: Press `Cmd + Space`, type `Terminal`, press Enter
- **Windows**: Press `Windows key`, type `cmd`, press Enter

### Step-by-step

Copy and paste each of these commands one at a time. Press **Enter** after each line.

#### Step 1: Go to the project folder

```bash
cd Desktop/application-tracker
```

> ℹ️ If you saved the folder somewhere else, use that path instead (e.g. `cd Downloads/application-tracker`)

#### Step 2: Install the backend (the server)

```bash
cd backend
npm install
```

> ℹ️ You'll see some text fly by — that's normal. Wait until it finishes and you see a new blinking line.

#### Step 3: Install the frontend (the web page)

```bash
cd ../frontend
npm install
```

> ℹ️ Same thing — wait for it to finish.

#### Step 4: Start the backend server

```bash
cd ../backend
npm start
```

> ℹ️ You should see: `Backend server listening at http://localhost:3000`
>
> **⚠️ Keep this terminal window open.** If you close it, the backend stops.

#### Step 5: Start the frontend (in a new terminal)

1. Open a **second** terminal window (same way as before)
2. Type this and press Enter:

```bash
cd Desktop/application-tracker/frontend
npm run dev
```

> ℹ️ You should see a link like `http://localhost:5173`

#### Step 6: Open the app

- Open your web browser (Chrome, Safari, Edge, etc.)
- Type or paste this into the address bar: **http://localhost:5173**
- Press Enter

🎉 **You're in!** The app should load with your dashboard.

---

## How to Use the App

The app has 4 tabs at the top of the page:

| Tab | What it does |
|-----|-------------|
| **Dashboard** | Shows an overview — upcoming deadlines, active applications, email stats |
| **Scholarships** | Add, edit, or delete scholarships you're tracking |
| **Applications** | Track university applications with a document checklist |
| **Email Tracker** | Log professor outreach emails and track responses |

Click the **"+ Add"** buttons to start entering your data. Everything saves automatically.

---

## How to Stop the App

- Go to each terminal window
- Press `Ctrl + C` (both Mac and Windows)
- Close the terminal windows

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "command not found: node" | Node.js isn't installed. Go back and download it from https://nodejs.org |
| "port 3000 already in use" | Close any other program using port 3000, or restart your computer |
| App loads blank / nothing shows | Make sure **both** terminal windows are running (one for backend, one for frontend) |
| "cannot find module..." | Run `npm install` again in both the `backend` and `frontend` folders |

---

## Built With

- **Backend**: Node.js, Express, SQLite (stores all your data in a file)
- **Frontend**: React, Vite (makes the interface fast and smooth)
