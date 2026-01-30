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
