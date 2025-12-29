# RogoAI Mail Hub v1

**Gmail Washing Machine Strategy: Secure Unified Management of Multiple Provider Emails**

[![YouTube Channel](https://img.shields.io/badge/YouTube-老後AI-red?logo=youtube)](https://www.youtube.com/@seniorAI-japan)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## 📖 Overview

RogoAI Mail Hub v1.0 is an email client that forwards emails from multiple ISP accounts (OCN, Nifty, So-net, Biglobe, etc.) to Gmail and manages them securely on your local machine.

### What is the Gmail Washing Machine Strategy?

```
ISP Email Accounts
    ↓ (Forward)
Gmail (Spam Filter = Washing Machine)
    ↓ (POP3 Retrieval)
Mail Hub (Local Storage = Security)
    ↓ (SMTP Send)
Reply from Original Address
```

1. **Forward ISP emails to Gmail**
   - Gmail's powerful spam filter "washes" your emails
   
2. **Retrieve with Mail Hub via POP3**
   - Only clean emails are downloaded locally
   - Safely stored in SQLite database
   
3. **Ransomware Protection**
   - Isolated from cloud services for local management
   - Access emails even offline
   
4. **Reply from Original Address**
   - Recipients see emails as sent from your original ISP address

---

## ✨ Key Features

- 📬 **Unified Multi-Address Management**: Displays forwarded emails sorted by original address
- 🎨 **Color-Coded by Provider**: Instantly identify which address received the email
- 📎 **Attachment Support**: Save and open email attachments
- 🔍 **Search Functionality**: Search by sender, subject, or body
- 💾 **Local Storage**: Email bodies saved in SQLite (offline access)
- ✉️ **Reply Function**: Reply from original address (via SMTP)
- 🌐 **HTML Display**: Properly renders HTML emails (using tkinterweb)

---

## 🚀 Quick Start

### For Windows (Recommended)

#### Executable Version (EXE)

1. Download the latest release from [Releases](https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite/releases)
   - `MailHub_v1.zip` (Japanese version)

**Note:** Version v13 is currently Japanese only. English version will be available in future releases.

2. Extract the ZIP and run `MailHub_v1.exe`

3. Enter settings on first launch
   - Gmail address
   - Gmail app password
   - Provider settings (forwarding source addresses)

#### Python Version

```bash
# Clone repository
git clone https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite.git
cd 251225_MailHub_GmailSQLite

# Run
python main.py
```

---

## ⚙️ Initial Setup

### 1. Get Gmail App Password

1. Go to [Google Account](https://myaccount.google.com/) → Security
2. **Enable 2-Step Verification** (required)
3. Generate an "App Password"
   - App: "Mail"
   - Device: "Windows Computer"
4. Note the **16-digit password**

### 2. Forward ISP Email to Gmail

Configure forwarding to your Gmail address in your ISP's management panel.

**Example: OCN Mail**
1. OCN My Page → Mail Settings
2. Forwarding Settings → Enter Gmail address
3. Select "Keep original email" (recommended)

### 3. Mail Hub Configuration

1. Launch Mail Hub
2. Open "Settings" tab
3. Enter **Gmail Settings**
   - Email Address: `your-email@gmail.com`
   - App Password: `xxxx xxxx xxxx xxxx` (16 digits)
   - POP Server: `pop.gmail.com`
   - POP Port: `995`
   - SMTP Server: `smtp.gmail.com`
   - SMTP Port: `587`

4. Add **Provider Settings**
   - Forwarding source email address (e.g., `yourname@ocn.ne.jp`)
   - Select display color

5. Click "Connection Test" to verify

6. Click "Save"

---

## 📁 Directory Structure

```
251225_MailHub_GmailSQLite/
├── main.py                      # Main program
├── icon.ico                     # Application icon
├── README.md                    # Japanese README
├── README_EN.md                 # This file (English)
├── LICENSE                      # MIT License
├── LICENSE_THIRD_PARTY.md       # Third-party licenses
├── CHANGELOG.md                 # Change log
├── .gitignore                   # Git exclusion settings
│
├── lib/                         # Bundled libraries
│   ├── tkinterweb/             # HTML display library (MIT)
│   └── tkinterweb_tkhtml/      # HTML display library dependency
│
├── config/                      # Config files (auto-generated, Git-excluded)
│   └── config.json             # User settings
│
└── storage/                     # Data storage (auto-generated, Git-excluded)
    └── emails.db               # Email database
```

**Note**: 
- `config/` and `storage/` contain personal information and are not included in Git
- Automatically generated on first launch

---

## 🛠️ Usage

### Retrieving Emails

1. Click "Fetch" button in the "Mail" tab
2. Downloads latest emails from Gmail via POP3
3. Displays emails color-coded by provider

### Viewing Emails

- Select an email from the list to display its body below
- HTML emails are properly formatted
- "Open Attachment" button appears if attachments exist

### Replying to Emails

1. Select the email you want to reply to
2. Click "Reply" button
3. Enter your message and send
4. **Sent from your original ISP address** (via SMTP)

### Search Function

- Search by sender, subject, or body
- Supports multiple keyword filtering

---

## ⚠️ Important Notes

### About POP3 After 2026

The current version (v13) uses POP3, and **Gmail's POP3 support is expected to continue**.

What will be discontinued in January 2026:
- ❌ Gmail Web UI feature to "fetch emails from third-party accounts via POP3"

What will continue:
- ✅ POP3 access to Gmail from external clients (like Mail Hub)

For details, see [Google's official announcement](https://support.google.com/mail/).

### Security

- **Keep your app password secure**
- `config.json` stores an encrypted password
- Never share this file with others

---

## 🐛 Troubleshooting

### Cannot Retrieve Emails

✅ **Check:**
- Gmail app password is correct
- 2-Step Verification is enabled
- POP is enabled in Gmail settings
  - Gmail Settings → Forwarding and POP/IMAP → "Enable POP"

### Provider Not Displayed

✅ **Check:**
- Provider is registered in Settings tab "Provider Settings"
- Gmail forwarding is configured correctly
- Gmail is actually receiving emails

### Cannot Open Attachments

✅ **Check:**
- Write permissions for temporary file location
- Antivirus software not blocking

### HTML Emails Not Displaying

✅ **Check:**
- `lib/tkinterweb/` folder exists
- Python version is 3.8 or higher

---

## 🔧 Advanced Customization

### Changing Config and Data Storage Locations

By default, Mail Hub saves data to `config/` and `storage/` folders in the same directory as main.py.

**Most users don't need to change this.** Only if you have special requirements, you can change the storage location using environment variables.

### Environment Variables

- `MAILHUB_CONFIG_DIR`: Location for config file (config.json)
- `MAILHUB_STORAGE_DIR`: Location for email database (emails.db)

### Usage

#### Start_MailHub.bat (Standard)

For normal users:
```batch
Start_MailHub.bat
```
- Uses default location (same folder as main.py)
- No configuration needed

#### Start_MailHub_Custom.bat (Advanced)

To specify custom storage locations:

1. Open `Start_MailHub_Custom.bat` in a text editor
2. Uncomment and set these lines:

```batch
set MAILHUB_CONFIG_DIR=D:\MySettings\MailHub
set MAILHUB_STORAGE_DIR=D:\MyData\MailHub
```

3. Run `Start_MailHub_Custom.bat`

### Important Notes

⚠️ **Avoid these locations**:
- Network drives (permission issues)
- Paths with non-ASCII characters
- System folders (permission errors)
- External drives (will fail if disconnected)

✅ **Recommended locations**:
- Local disk with ASCII-only paths
- Folders with write permissions

---

## 🔧 Developer Information

### Requirements

- Python 3.8 or higher
- Standard library only (no additional installation required)

### Bundled Libraries

This project bundles the following libraries in the `lib/` folder:

- **tkinterweb**: For HTML email display (MIT License)
- **tkinterweb_tkhtml**: Dependency for tkinterweb

See `LICENSE_THIRD_PARTY.md` for details.

### Development Setup

```bash
# Clone repository
git clone https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite.git
cd 251225_MailHub_GmailSQLite

# Run
python main.py
```

### Building EXE

```bash
# Install PyInstaller
pip install pyinstaller

# Run build script
build_exe.bat
```

---

## 📝 License

MIT License - See [LICENSE](LICENSE) for details.

Third-party library licenses are in [LICENSE_THIRD_PARTY.md](LICENSE_THIRD_PARTY.md).

---

## 👤 Author

**Created by Takejii (RogoAI)**

- YouTube: [@seniorAI-japan](https://www.youtube.com/@seniorAI-japan)
- GitHub: [RogoAI-Takeji](https://github.com/RogoAI-Takeji)

---

## 🙏 Support

For questions or issues:
1. [GitHub Issues](https://github.com/RogoAI-Takejij/251225_MailHub_GmailSQLite/issues)
2. [YouTube Video Comments](https://www.youtube.com/@seniorAI-japan)

---

## 📚 Related Videos

- [Mail Hub Introduction](https://youtube.com/...)
- [Gmail Washing Machine Strategy Explained](https://youtube.com/...)
- [Importance of Security Measures](https://youtube.com/...)

---

**Enjoy safe email management! 📧🔒**