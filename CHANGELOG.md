# Changelog

All notable changes to RogoAI Mail Hub will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.3] - 2026-01-06

### Fixed
- **Settings Persistence (Critical)**: Fixed issue where settings were not saved after application restart in exe mode
  - Settings now correctly save to `config/` folder next to the executable
  - Database now saves to `storage/` folder next to the executable
  - Added `sys.frozen` detection to use `sys.executable` path for exe mode
- **SMTP Connection Error**: Fixed `[SSL:WRONG_VERSION_NUMBER]` error when using port 587
  - Added support for both port 587 (STARTTLS) and port 465 (SSL/TLS)
  - Automatically detects port type and uses appropriate connection method
  - STARTTLS: `smtplib.SMTP()` + `starttls()`
  - SSL/TLS: `smtplib.SMTP_SSL()`
- **UI Responsiveness**: Fixed application freeze during email sending
  - SMTP operations now run in separate threads
  - "Sending..." indicator displayed during send operation
  - Send button disabled during operation to prevent multiple clicks
  - Applied to both compose and reply functions
- **Security Enhancement**: Changed default email display mode
  - Emails now open in text mode by default (safer)
  - HTML display requires explicit button click ("HTML (Safe)" or "HTML (Full)")
  - Prevents unintended HTML content execution

### Changed
- Improved user experience with threaded email operations
- Enhanced security posture with text-first display policy
- Build script now includes `--hidden-import tkinterweb` for stable HTML display

### Technical Details
- Modified BASE_DIR calculation to support exe mode path resolution
- Implemented threading for SMTP send operations in:
  - `open_compose_window` do_send function
  - `open_reply_window` do_send function
  - `open_draft_editor` do_send function
- Added port-based SMTP connection logic (465 vs 587)
- Modified `MailViewer.__init__` to default to `switch_mode("text")` instead of `switch_mode("html_safe")`

---

## [1.0.1] - 2026-01-04

### Fixed
- **Context Menu Bug**: Fixed issue where right-click "Move to Folder" submenu would not display custom folders for emails in Promo Box
- **TypeError Fix**: Resolved `TypeError: unhashable type: 'list'` that prevented right-click context menu from appearing
- **Provider Detection**: Improved actual provider detection for promotional emails by extracting domain from sender's email address
- **Folder Menu Display**: Custom folders now correctly appear in both single-selection and multi-selection context menus

### Changed
- Enhanced `build_move_menu_single()` and `build_move_menu_multiple()` methods to properly handle Promo Box emails
- Added error handling and fallback mechanisms to ensure context menu always displays even when errors occur

---

## [1.0.0] - 2025-12-28

### Added
- **Gmail Washing Machine Strategy**: Forward ISP emails to Gmail, filter through spam detection, and manage locally
- **IMAP Email Retrieval**: Secure email download from Gmail via IMAP
- **SMTP Reply Function**: Reply from original ISP address via SMTP (supports both port 587 and 465)
- **Provider Management**: Add multiple ISP email addresses with color coding
- **SQLite Database**: Local email storage for offline access
- **HTML Email Display**: Proper rendering of HTML emails using tkinterweb with safe/full modes
- **Attachment Support**: Save and open email attachments
- **Search Functionality**: Search emails by sender, subject, or body
- **Connection Test**: Verify Gmail settings before saving
- **Japanese Interface**: All GUI elements in Japanese (English version planned for future releases)

### Changed
- Enhanced GUI layout with tkinter
- Improved security with encrypted password storage (Base64)

### Security
- **Ransomware Protection**: Emails stored locally, isolated from cloud
- **Encrypted Credentials**: App passwords stored with Base64 encoding
- **Local-only Storage**: config.json and emails.db excluded from version control
- **Safe HTML Display**: Default text mode with explicit HTML activation

### Technical Details
- **Protocol**: IMAP for retrieval, SMTP (port 587 or 465) for sending
- **Database**: SQLite for email storage
- **GUI Framework**: Tkinter with tkinterweb for HTML rendering
- **Dependencies**: Bundled tkinterweb and tkinterweb_tkhtml (no additional installation required)
- **Build System**: PyInstaller with `--onefile` flag for single executable distribution

---

## Notes

### Version History Summary

- **v1.0.3**: Critical bug fixes (settings, SMTP, threading, security)
- **v1.0.1**: Context menu bug fixes
- **v1.0.0**: Initial release

### Known Issues

None reported in v1.0.3.

### Future Enhancements

- OAuth2 support for enhanced security
- English UI localization
- Advanced filtering options
- Email template system

---

## Support

For questions or bug reports:
- **GitHub Issues**: https://github.com/RogoAI-Takeji/251224_RogoAI_Mail_Hub/issues
- **YouTube**: https://www.youtube.com/@seniorAI-japan

---

**Version Naming**: Version numbers follow semantic versioning (MAJOR.MINOR.PATCH) for better tracking of changes and compatibility.
