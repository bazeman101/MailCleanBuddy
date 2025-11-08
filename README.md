# ✉️ MailCleanBuddy - Interactive Mailbox Manager for Microsoft 365

<div align="center">

[![Version](https://img.shields.io/badge/version-3.1-blue.svg)](https://github.com/bazeman101/MailCleanBuddy/releases/tag/v3.1)
[![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Production Ready](https://img.shields.io/badge/status-production%20ready-success.svg)](https://github.com/bazeman101/MailCleanBuddy)

### ☕ Support This Project

**If MailCleanBuddy helps you manage your inbox, consider supporting development:**

<a href="https://buymeacoffee.com/basw" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" ></a>

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-Support%20Development-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/basw)

**Your support helps keep this project free and actively maintained!** 🙏

</div>

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [What's New in v3.1](#-whats-new-in-v31)
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Usage Guide](#-usage-guide)
- [System Requirements](#-system-requirements)
- [Configuration](#️-configuration)
- [Troubleshooting](#-troubleshooting)
- [Contributing](#-contributing)
- [License](#-license)
- [Support](#-support)

---

## 🎯 Overview

**MailCleanBuddy** is a powerful, interactive PowerShell tool for managing Microsoft 365 mailboxes. With 27 modular components, advanced search capabilities, threat detection, and smart automation, it helps you keep your inbox organized and secure.

### ✨ Key Highlights

- 🎨 **Modern UI** - Intuitive menu-driven interface with color-coded actions
- 🔒 **Security First** - Multi-layer threat detection (phishing, malware, spoofing)
- 📊 **Smart Analytics** - Comprehensive insights into your email patterns
- 🌍 **Multi-Language** - Full support for Dutch, English, German, and French
- ⚡ **High Performance** - Local caching for instant search and navigation
- 🛠️ **Modular Design** - 27 independent modules for clean architecture

---

## 🚀 Features

### 📧 Email Management
- ✅ **Advanced Search** with regex support and exact phrase matching
- ✅ **Arrow Key Navigation** (← previous, → next email)
- ✅ **Bulk Operations** (delete, move, archive multiple emails)
- ✅ **Smart Folder Organization** with auto-learning
- ✅ **Email Export** to EML/MSG format
- ✅ **HTML Browser Preview** for emails

### 🔒 Security & Threat Detection
- ✅ **Multi-Layer Threat Detection** (phishing, malware, spoofing)
- ✅ **Intelligent Scoring System** for suspicious emails
- ✅ **Quarantine Management** for risky messages
- ✅ **DKIM/SPF/DMARC Analysis** with header inspection
- ✅ **Link Safety Checking** for dangerous URLs
- ✅ **Suspicious Attachment Detection**

### 📊 Analytics & Insights
- ✅ **Comprehensive Dashboard** with mailbox statistics
- ✅ **Attachment Statistics** with 4-tier fallback calculation
- ✅ **Storage Usage Analysis** by sender and domain
- ✅ **Growth Trends** visualization
- ✅ **Large Attachment Manager** with download/delete options
- ✅ **Duplicate Email Detection** (quick & deep scan modes)

### ⚙️ Advanced Features
- ✅ **VIP Sender Management** with protection from accidental deletion
- ✅ **Thread/Conversation Analysis** with bulk actions
- ✅ **Unsubscribe Manager** for newsletters with confidence scoring
- ✅ **Email Archiving** with retention policies
- ✅ **Calendar Sync** capabilities
- ✅ **Custom Rules** and automation

### 🌍 Internationalization
- ✅ **4 Languages**: Dutch (nl), English (en), German (de), French (fr)
- ✅ **690+ Localized Strings** for complete translation
- ✅ **Dynamic Language Switching** at runtime
- ✅ **Culture-Aware Formatting** for dates and numbers

### 🏗️ Technical Excellence
- ✅ **27 Modular Components** with clean separation
- ✅ **Robust Error Handling** with verbose logging
- ✅ **Local Cache System** for performance
- ✅ **Microsoft Graph API** integration
- ✅ **Input Validation** for all parameters
- ✅ **Production Ready** (Quality Score: 85/100)

---

## 🎉 What's New in v3.1

**Release Date:** 2025-11-08
**Status:** ✅ Production Ready

### Critical Improvements

#### 🛡️ ColorScheme Null-Safety
- Prevents crashes if ColorScheme module fails to load
- Automatic fallback initialization
- Safe color property access with `Get-SafeColor` function

#### ✅ Enhanced Input Validation
- Email address format validation with regex
- Language selection restricted to supported values
- Range validation for numeric parameters
- Prevents invalid input from causing errors

#### 🔍 Verbose Error Logging
- Silent catch blocks now log to `Write-Verbose`
- Better diagnostics for MAPI property parse failures
- Improved debugging for developers

#### 🌐 Complete Dutch Localization
- All English strings translated to Dutch
- Consistent language experience
- No more mixed language in UI

#### 🎯 Clean Output
- Eliminated unwanted "True" output after operations
- All Graph API calls properly suppress return values
- Professional user experience

### Quality Metrics

| Metric | v3.0 | v3.1 | Improvement |
|--------|------|------|-------------|
| Error Resilience | 65/100 | 75/100 | +10 ✅ |
| Localization Quality | 80/100 | 95/100 | +15 ✅ |
| Code Maintainability | 85/100 | 85/100 | - |
| Documentation | 90/100 | 90/100 | - |
| **Overall Production Readiness** | 75/100 | **85/100** | **+10 ✅** |

**Full Release Notes:** [release-notes/RELEASE_NOTES_v3.1.md](release-notes/RELEASE_NOTES_v3.1.md)

---

<div align="center">

## ☕ Enjoying MailCleanBuddy?

**If this tool saves you time and makes email management easier, consider supporting its development!**

<a href="https://buymeacoffee.com/basw" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" ></a>

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/basw)

**Your support helps:**
- 🔧 Continue active development
- ✨ Add new features
- 🐛 Fix bugs quickly
- 📚 Improve documentation
- 💚 Keep the project free and open-source

**Thank you for your support! 🙏**

</div>

---

## 📥 Installation

### Prerequisites

- **PowerShell 7+** (recommended) or **Windows PowerShell 5.1**
- **Microsoft 365 Account** with mailbox access
- **Internet Connection** for Microsoft Graph API

### Method 1: Git Clone (Recommended)

```powershell
# Clone the repository
git clone https://github.com/bazeman101/MailCleanBuddy.git
cd MailCleanBuddy

# Run the script
.\MailCleanBuddy.ps1 -MailboxEmail "your@email.com"
```

### Method 2: Download ZIP

1. Download the [latest release](https://github.com/bazeman101/MailCleanBuddy/releases/latest)
2. Extract to your preferred location
3. Open PowerShell in the extracted folder
4. Run: `.\MailCleanBuddy.ps1 -MailboxEmail "your@email.com"`

### First-Time Setup

On first run, the script will:
1. ✅ Check for required PowerShell modules
2. ✅ Install missing modules (`Microsoft.Graph.Authentication`, `Microsoft.Graph.Mail`)
3. ✅ Connect to Microsoft Graph (browser authentication)
4. ✅ Build local email cache (may take time for large mailboxes)

**Note:** Module installation requires administrator privileges on Windows.

---

## 🚀 Quick Start

### Basic Usage

```powershell
# Dutch interface (default)
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com"

# English interface
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Language en

# German interface
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Language de

# French interface
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Language fr
```

### Test Mode

```powershell
# Index only 100 most recent emails (fast testing)
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -TestRun

# Index only 1000 most recent emails
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -MaxEmailsToIndex 1000
```

### Advanced Examples

```powershell
# Verbose logging for troubleshooting
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Verbose

# English + test mode + verbose
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Language en -TestRun -Verbose
```

---

## 📖 Usage Guide

### Main Menu Structure

```
╔══════════════════════════════════════════════════════════╗
║          MailCleanBuddy - Main Menu                      ║
╠══════════════════════════════════════════════════════════╣
║  1. View inbox                                           ║
║  2. Manage emails from specific sender                   ║
║  3. Advanced email search                                ║
║  4. View recent emails (last 100)                        ║
║  5. Bulk attachment download                             ║
║  6. Delete old emails                                    ║
║  7. Smart features (Analytics, Security, Automation)     ║
║  8. Rebuild cache                                        ║
║  9. Language settings                                    ║
║  Q. Quit                                                 ║
╚══════════════════════════════════════════════════════════╝
```

### Email Viewer Shortcuts

When viewing an email:

| Key | Action |
|-----|--------|
| **B** | View body/content |
| **O** | Open in browser (HTML preview) |
| **H** | View email headers |
| **D** | Download attachments |
| **Delete** | Delete email |
| **V** | Move to folder |
| **←** | Previous email |
| **→** | Next email |
| **Esc** | Back to list |

### Email List Shortcuts

When browsing emails:

| Key | Action |
|-----|--------|
| **↑/↓** | Navigate emails |
| **Space** | Select multiple emails |
| **A** | Select/deselect all |
| **Delete** | Delete selected |
| **V** | Move selected |
| **Enter** | Open email |
| **Esc** | Back to menu |

### Advanced Search

```
Search Options:
- Regex search: /pattern/
  Example: /invoice-\d{4}/ finds "invoice-2024"

- Exact phrase: "phrase in quotes"
  Example: "urgent action required"

- Quick filters:
  [A] Emails WITH attachments
  [N] Emails WITHOUT attachments
  [S] Search by specific sender
```

---

## 💻 System Requirements

### Operating Systems
- ✅ Windows 10/11
- ✅ Windows Server 2016+
- ✅ Linux (with PowerShell 7+)
- ✅ macOS (with PowerShell 7+)

### PowerShell
- **PowerShell 7+** (recommended) - [Download here](https://github.com/PowerShell/PowerShell/releases)
- **Windows PowerShell 5.1** (compatible, but limited features)

### Required Modules
Automatically installed on first run:
- `Microsoft.Graph.Authentication` - For OAuth authentication
- `Microsoft.Graph.Mail` - For email operations

### Microsoft Graph Permissions
Required scopes (auto-requested):
- `Mail.Read` - Read email
- `Mail.ReadWrite` - Manage email (delete, move, etc.)

### Hardware Recommendations
- **RAM:** 4GB minimum, 8GB+ recommended
- **Storage:** 100MB for application + cache storage
- **Network:** Stable internet connection for Graph API calls

---

## ⚙️ Configuration

### Command-Line Parameters

```powershell
.\MailCleanBuddy.ps1 [parameters]

Parameters:
  -MailboxEmail <string>     (Required) Email address to manage
  -Language <string>         UI language: nl, en, de, fr (default: nl)
  -TestRun                   Index only 100 most recent emails
  -MaxEmailsToIndex <int>    Limit indexing to N emails (0 = all)
  -Verbose                   Enable detailed logging
```

### Cache Location

Cache files are stored in:
```
<script-directory>/cache_<email-address>.json
```

Example: `C:\Tools\MailCleanBuddy\cache_user@example.com.json`

### Language Files

Localizations are stored in:
```
<script-directory>/localizations.json
```

Contains translations for: Dutch, English, German, French

---

## 🔧 Troubleshooting

### Common Issues

#### ❌ "Cannot connect to Microsoft Graph"

**Solution:**
1. Ensure you have internet connectivity
2. Check firewall/proxy settings
3. Try running: `Connect-MgGraph -Scopes "Mail.Read","Mail.ReadWrite"`
4. Clear token cache: `Disconnect-MgGraph` then reconnect

#### ❌ "Module Microsoft.Graph.Mail not found"

**Solution:**
```powershell
# Install manually with admin privileges
Install-Module Microsoft.Graph.Authentication -Force
Install-Module Microsoft.Graph.Mail -Force
```

#### ❌ "Execution Policy Error"

**Solution:**
```powershell
# Set execution policy (run as Administrator)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### ❌ "Cache loading is slow"

**Causes:**
- Large mailbox (10,000+ emails)
- First-time indexing

**Solutions:**
- Use `-TestRun` for initial testing
- Use `-MaxEmailsToIndex 1000` to limit scope
- Wait for cache to build (one-time process)
- Subsequent runs will be fast (loads from cache)

#### ❌ "Search returns no results"

**Solutions:**
- Rebuild cache: Menu option 8
- Check search syntax (regex needs `/pattern/`)
- Verify emails exist in indexed range
- Try broader search terms

### Verbose Logging

Enable detailed logging for troubleshooting:

```powershell
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Verbose
```

### Reset Everything

```powershell
# Delete cache file
Remove-Item "cache_user@example.com.json"

# Disconnect Graph session
Disconnect-MgGraph

# Run script again
.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com"
```

---

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

### Reporting Bugs

1. Check [existing issues](https://github.com/bazeman101/MailCleanBuddy/issues)
2. Create a new issue with:
   - Clear description
   - Steps to reproduce
   - Expected vs actual behavior
   - PowerShell version
   - Error messages (if any)

### Suggesting Features

1. Check [roadmap](FEATURES_ROADMAP.md) for planned features
2. Create an issue with `enhancement` label
3. Describe the feature and use case

### Development Setup

```powershell
# Clone repository
git clone https://github.com/bazeman101/MailCleanBuddy.git
cd MailCleanBuddy

# Run tests (in dev-tools folder)
.\dev-tools\Test-AllModules-Parser.ps1
```

### Pull Requests

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test thoroughly
5. Commit (`git commit -m 'Add amazing feature'`)
6. Push (`git push origin feature/amazing-feature`)
7. Open a Pull Request

---

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2025 bazeman101

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
```

---

## 🔗 Links

- **GitHub Repository:** https://github.com/bazeman101/MailCleanBuddy
- **Latest Release:** https://github.com/bazeman101/MailCleanBuddy/releases/latest
- **Report Issues:** https://github.com/bazeman101/MailCleanBuddy/issues
- **Discussions:** https://github.com/bazeman101/MailCleanBuddy/discussions
- **Release Notes:** [release-notes/](release-notes/)
- **Roadmap:** [FEATURES_ROADMAP.md](FEATURES_ROADMAP.md)

---

## 💪 Support

### Get Help

- 📖 **Documentation:** Read this README and [release notes](release-notes/)
- 🐛 **Bug Reports:** [Create an issue](https://github.com/bazeman101/MailCleanBuddy/issues)
- 💬 **Discussions:** [Ask questions](https://github.com/bazeman101/MailCleanBuddy/discussions)
- 📧 **Email:** Contact via GitHub

### Support Development

<div align="center">

**If MailCleanBuddy makes your email management easier, consider buying me a coffee!** ☕

<a href="https://buymeacoffee.com/basw" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" ></a>

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-Support%20Development-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/basw)

**Your support helps:**
- ✨ Add new features
- 🐛 Fix bugs faster
- 📚 Improve documentation
- 🔒 Enhance security
- 💚 Keep the project free and open-source

</div>

---

## 🙏 Acknowledgments

- **Microsoft Graph API Team** - For excellent API documentation
- **PowerShell Community** - For amazing modules and support
- **All Contributors** - Thank you for your feedback and contributions!
- **You!** - For using MailCleanBuddy and supporting its development

---

<div align="center">

**Made with ❤️ by [bazeman101](https://github.com/bazeman101)**

**Version 3.1** | **Production Ready** | **MIT Licensed**

[![Star on GitHub](https://img.shields.io/github/stars/bazeman101/MailCleanBuddy?style=social)](https://github.com/bazeman101/MailCleanBuddy/stargazers)
[![Fork on GitHub](https://img.shields.io/github/forks/bazeman101/MailCleanBuddy?style=social)](https://github.com/bazeman101/MailCleanBuddy/network/members)

</div>

---

<div align="center">

### ⭐ If you find this project useful, please give it a star! ⭐

**Star History**

[![Star History Chart](https://api.star-history.com/svg?repos=bazeman101/MailCleanBuddy&type=Date)](https://star-history.com/#bazeman101/MailCleanBuddy&Date)

</div>

---

*Happy email managing! 📧✨*
