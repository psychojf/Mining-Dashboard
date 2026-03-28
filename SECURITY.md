# Security Policy

## Supported Versions

This project is currently in active development. Security updates are provided for the following versions:

| Version | Supported          |
| ------- | ------------------ |
| Latest  | :white_check_mark: |
| Older   | :x:                |

**Note:** Always use the latest version from the `main` branch to ensure you have the most recent security patches and updates.

## Security Considerations

### Application Design

This application is designed as a **local, single-user tool** that:

* Runs entirely on your local machine
* Reads EVE Online game logs from your local filesystem
* Does NOT transmit sensitive data over the internet (except optional Discord webhooks)
* Does NOT require authentication or store passwords
* Does NOT access EVE Online's servers directly
* Does NOT modify game files or interact with the game client

### Data Privacy

The mining dashboard:

* **Game Logs:** Only reads EVE Online combat/mining logs stored locally on your system
* **Configuration:** Stores settings in local JSON files in the application directory
* **Discord Webhooks (Optional):** If enabled, sends mining statistics to your configured Discord webhook URL
* **No Telemetry:** Does NOT collect, transmit, or store any personal information or usage statistics
* **SDE Data:** Downloads CCP's official Static Data Export (publicly available JSONL files) for ore/ice/gas information

### Executable Security

When using the PyInstaller-compiled executable:

* **Source Verification:** Always download releases from the official GitHub repository
* **Antivirus False Positives:** PyInstaller executables may trigger false positives in some antivirus software. This is a known limitation of Python-to-executable packaging.
* **Build Verification:** If concerned, you can build the executable yourself from source using the provided requirements and PyInstaller configuration

### Discord Webhook Security

If you choose to enable Discord integration:

* **Webhook URLs are sensitive:** Treat your Discord webhook URL like a password
* **Never share your webhook URL publicly** - anyone with the URL can send messages to your Discord channel
* **Webhook data:** Only mining statistics (ore quantities, character names, session info) are sent - no API keys or credentials
* **HTTPS Only:** All webhook communications use HTTPS encryption
* **Revoke Compromised Webhooks:** If your webhook URL is exposed, regenerate it in Discord immediately

## Reporting a Vulnerability

### What to Report

Please report security vulnerabilities if you discover:

* **Code injection vulnerabilities** in log parsing or file handling
* **Path traversal attacks** that could access files outside the intended directories
* **Dependency vulnerabilities** in third-party Python libraries
* **Data leakage** that exposes sensitive information unintentionally
* **Privilege escalation** issues
* **Discord webhook abuse vectors**
* **Malicious log files** that could exploit the parser

### What NOT to Report

The following are **not security vulnerabilities**:

* Features that require manual user configuration (Discord webhooks, file paths)
* EVE Online gameplay mechanics or balance issues
* CCP Games server-side security (report those to CCP directly)
* Issues with CCP's Static Data Export files
* Antivirus false positives on the PyInstaller executable

### How to Report

**DO NOT open a public GitHub issue for security vulnerabilities.**

Instead, please:

1. **Open a GitHub Security Advisory** (preferred):
   * Go to the repository's Security tab
   * Click "Report a vulnerability"
   * Provide detailed information about the vulnerability

2. **Contact via GitHub Issues** (for non-critical issues):
   * Open a private issue describing the concern
   * Tag it with `security` label if available

### What to Include in Your Report

Please provide:

* **Description:** Clear explanation of the vulnerability
* **Steps to Reproduce:** Detailed steps to demonstrate the issue
* **Impact:** What an attacker could accomplish
* **Affected Versions:** Which versions are vulnerable
* **Suggested Fix:** If you have a proposed solution (optional)
* **Environment:** Python version, OS, relevant configuration

Example:
```
Title: Path Traversal in Log File Selection

Description: The application allows users to select log files without 
properly validating the file path, potentially allowing access to 
files outside the EVE Online logs directory.

Steps to Reproduce:
1. Configure log path to include "../../../"
2. Observe that files outside intended directory are accessible

Impact: An attacker could potentially read arbitrary files on the 
user's system if they can control the log path configuration.

Affected Versions: All versions up to commit [hash]

Suggested Fix: Implement path validation using os.path.normpath() 
and verify the resolved path is within the allowed directory.
```

## Response Timeline

* **Initial Response:** Within 72 hours of report submission
* **Vulnerability Assessment:** Within 7 days
* **Patch Development:** Varies based on severity and complexity
* **Public Disclosure:** After a fix is released and users have had time to update (typically 30 days)

## Security Best Practices for Users

### When Installing

1. **Download from official sources only** (GitHub releases)
2. **Verify file integrity** if checksums are provided
3. **Scan with antivirus** (expect possible false positives with PyInstaller)
4. **Review permissions** - the app should only need read access to EVE logs directory

### When Using

1. **Keep the application updated** to the latest version
2. **Protect your Discord webhook URLs** - never commit them to git or share publicly
3. **Run with minimal privileges** - no administrator/root access required
4. **Review log files** before processing if received from untrusted sources
5. **Use firewall rules** if you want to restrict network access (only Discord webhooks need internet)

### When Building from Source

1. **Install dependencies in a virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. **Review dependencies for known vulnerabilities**
   ```bash
   pip install safety
   safety check
   ```

3. **Keep Python and dependencies updated**

## Dependency Security

This project uses the following external dependencies:

* **tkinter** - Standard library (bundled with Python)
* **openpyxl** - Excel file generation
* **plyer** - Desktop notifications
* **requests** - HTTP library for Discord webhooks and SDE downloads
* **winsound** - Windows audio (standard library)

Dependencies are tracked in `requirements.txt`. We monitor for security advisories and update dependencies when vulnerabilities are discovered.

## CCP Games EULA Compliance

This tool is designed to comply with CCP Games' End User License Agreement and Terms of Service:

* **Read-Only Access:** Only reads publicly documented log files
* **No Game Modification:** Does not modify game files or memory
* **No Automation:** Does not automate gameplay or interact with the game client
* **No Server Interaction:** Does not communicate with EVE Online servers
* **Informational Only:** Provides statistics based on local log data

**Using third-party tools is at your own risk.** While this tool is designed to be EULA-compliant, always review CCP's current policies.

## License

This project is open source under [specify your license]. Security fixes will be made available under the same license.

## Acknowledgments

We appreciate responsible disclosure of security vulnerabilities. Contributors who report valid security issues will be acknowledged in release notes (unless they prefer to remain anonymous).

---

**Last Updated:** March 2026

*Fly safe, mine smart, stay secure.* 🔒⛏️
