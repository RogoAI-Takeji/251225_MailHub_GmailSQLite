# Changelog

All notable changes to RogoAI Mail Hub will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] - 2025-01-04

### Added
- **Gmail Washing Machine Strategy**: Forward ISP emails to Gmail, filter through spam detection, and manage locally
- **POP3 Email Retrieval**: Secure email download from Gmail via POP3_SSL
- **SMTP Reply Function**: Reply from original ISP address via SMTP
- **Provider Management**: Add multiple ISP email addresses with color coding
- **SQLite Database**: Local email storage for offline access
- **HTML Email Display**: Proper rendering of HTML emails using tkinterweb
- **Attachment Support**: Save and open email attachments
- **Search Functionality**: Search emails by sender, subject, or body
- **Connection Test**: Verify Gmail settings before saving
- **Japanese Interface**: All GUI elements in Japanese (English version planned for future releases)

### Changed
- Transitioned from IMAP to POP3 for better ransomware protection
- Improved security with encrypted password storage (Base64)
- Enhanced GUI layout with tkinter

### Security
- **Ransomware Protection**: Emails stored locally, isolated from cloud
- **Encrypted Credentials**: App passwords stored with Base64 encoding
- **Local-only Storage**: config.json and emails.db excluded from version control

### Technical Details
- **Protocol**: POP3_SSL (port 995) for retrieval, SMTP (port 587) for sending
- **Database**: SQLite for email storage
- **GUI Framework**: Tkinter with tkinterweb for HTML rendering
- **Dependencies**: Bundled tkinterweb and tkinterweb_tkhtml (no additional installation required)

---

## Notes

### Future Considerations

- **2026 POP3 Continuity**: Gmail's POP3 access from external clients is expected to continue beyond 2026
- **IMAP Migration**: If Google discontinues POP3 entirely, IMAP migration will be considered
- **OAuth2 Support**: May be added in future versions for enhanced security

### Known Issues

None reported in this release.

---

## Support

For questions or bug reports:
- **GitHub Issues**: https://github.com/yourusername/251225_MailHub_GmailSQLite/issues
- **YouTube**: https://www.youtube.com/@rogoai

---

**Version Naming**: v1.0.0 corresponds to the video release date (2025/01/04) and represents the initial public release.