# VideoAsk Automation Bot (Selenium)

This project automates sending VideoAsk messages using Selenium.

---

## ✅ Requirements

* Python **3.10+** (Recommended)
* Google Chrome installed

---

## 📦 Installation

### 1) Install Python

Download and install Python from the official website:

[https://www.python.org/downloads/](https://www.python.org/downloads/)

> During installation, make sure **“Add Python to PATH”** is checked.

---

### 2) Install Dependencies

Run:

```bash
python install_dependencies.py
```

---

## ⚙️ Configuration (`config.ini`)

Open `config.ini` and update the following values:

* `login` → your username/email
* `password` → your password
* `title` → message title
* `button_text` → button text
* `button_url` → button link

### Skip Rules

* `skip_if_message_already_sent` = **yes / no**

  * **yes** → Skip lead if message was already sent to that lead
  * **no** → Always send message again

* `skip_if_message_exists` = **yes / no**

  * **yes** → Skip lead if a message with the same title already exists
  * **no** → Always send message

### Leads File

* `contacts_file` = **CSV file name or full file path**

  * Leave empty → Bot will send messages to all contacts
  * Provide CSV file → Bot will send messages only to those leads

### Bot Speed

* `speed` = **slow / normal / fast**

---

## ▶️ Run the Bot

Start the automation:

```bash
python "main v2.py"
```

---

## ⚠️ Important Notes

* The bot opens Chrome automatically.
* Keep the Chrome window **wide enough** so all elements remain visible.
* Do not minimize Chrome while the bot is running.

---

# 🚀 Publish / Push to GitHub (Windows CMD)

## 1) Open CMD / PowerShell and go to your project folder

```bat
cd /d "E:\Learning Python\videoask"
```

---

## 2) Fix “dubious ownership” (after Windows reinstall)

Run this (CMD/PowerShell compatible):

```bat
git config --global --add safe.directory "E:/Learning Python/videoask"
```

---

## 3) Initialize git (if needed)

```bat
git init
```

---

## 4) Add files and commit

```bat
git add .
git commit -m "Initial commit"
```

---

## 5) Connect to GitHub repo

```bat
git remote remove origin
git remote add origin https://github.com/bilal-virk/videoask.git
```

---

## 6) Push to GitHub

```bat
git branch -M main
git push -u origin main
```

---

## 🛠️ If you still get ownership issues (Permanent Fix)

Run CMD as **Administrator**:

```bat
takeown /F "E:\Learning Python\videoask" /R /D Y
icacls "E:\Learning Python\videoask" /grant "%USERNAME%":F /T
```
