
# 🛠️ PowerShell Administrative Scripts Collection

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-GPL%20v3-green.svg)](LICENSE)
[![Last Update](https://img.shields.io/badge/Last%20Update-October%202025-brightgreen)](https://github.com/roalhelm/PowershellScripts)
[![Scripts](https://img.shields.io/badge/Scripts-25%2B-orange)](https://github.com/roalhelm/PowershellScripts)

A comprehensive collection of PowerShell scripts for system administration, Microsoft Intune, Windows Updates, user/group management, network diagnostics, and remediation tasks in modern Windows enterprise environments.

## 🌟 Highlights

- **🔌 Microsoft Endpoint Connectivity Tester V2.1** - Advanced connectivity, latency, and performance tests with HTML reports

---

## 🚀 Quick Start

```powershell
# Clone repository
git clone https://github.com/roalhelm/PowershellScripts.git
cd PowershellScripts

# Example: Microsoft Endpoint Connectivity Test with HTML Report
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport -OpenReport

# Example: Test only Intune and Defender
.\CheckMicrosoftEndpointsV2.ps1 -Services Intune,Defender -HtmlReport "MyReport.html"
```

---

## 📋 System Requirements

| Requirement | Details |
|-------------|---------|
| **PowerShell** | 5.1 or higher |
| **Permissions** | Administrator rights for most scripts |
| **Operating System** | Windows 10/11, Windows Server 2016+ |
| **Modules** | PSWindowsUpdate, Microsoft.Graph, ActiveDirectory |

---

## 🏆 Featured Script: Microsoft Endpoint Connectivity Tester V2.1

### ✨ New Features in Version 2.1
- **🎨 HTML Report Generation** - Professional, responsive reports
- **🎯 Service Selection** - Test only relevant Microsoft Services
- **⚡ Performance Options** - Skip ping/speed tests for faster execution
- **📱 Responsive Design** - Reports work on all devices
- **🔍 Interactive Menu** - User-friendly service selection

### 🎮 Usage

```powershell
# Interactive mode (recommended for new users)
.\CheckMicrosoftEndpointsV2.ps1

# All services with full report
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport -OpenReport

# Quick test for critical services only
.\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,Intune,AzureAD -SkipSpeed

# Automated test for CI/CD
.\CheckMicrosoftEndpointsV2.ps1 -Services All -Quiet -HtmlReport "report.html"
```

### 🎯 Supported Microsoft Services
- **Windows Update for Business** - Update endpoints and delivery services
- **Windows Autopatch** - Automatic patch management
- **Microsoft Intune** - Device management and compliance
- **Microsoft Defender** - Security and threat protection
- **Azure Active Directory** - Identity services and authentication
- **Microsoft 365** - Office apps and cloud services
- **Microsoft Store** - App distribution and updates
- **Windows Activation** - Licensing and activation
- **Microsoft Edge** - Browser services and enterprise features
- **Windows Telemetry** - Diagnostic data and error reporting

---

## 🎨 HTML Report Features (CheckMicrosoftEndpointsV2.ps1)

### 📊 Dashboard Overview
- **Statistical Cards** - Tested endpoints, success/failure rate, performance metrics
- **Color-coded Indicators** - Instant visual assessment
- **Responsive Grid Layout** - Works on desktop, tablet, mobile

### 📋 Detailed Service Tables
- **Service-specific Grouping** - Clear organization by Microsoft Services
- **Status Badges** - OK/FAILED with color coding
- **IP Addresses** - For network troubleshooting
- **Performance Metrics** - Latency and speed data (optional)

### 📈 Performance Analysis
- **Latency Rating** - Automatic classification (Excellent/Good/Needs Improvement)
- **Speed Statistics** - Min/Max/Average
- **Service Impact Analysis** - What failures mean in practice

### 🎨 Modern Design
- **Microsoft Design Language** - Familiar look for IT professionals
- **Gradient Headers** - Professional appearance
- **Shadows and Animations** - Modern web aesthetics
- **Print-friendly** - Optimized for PDF export

---

## 💼 Practical Examples

### 🔧 Daily IT Administration
```powershell
# Morning network check with report
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "Daily-$(Get-Date -Format 'yyyy-MM-dd').html" -OpenReport

```

### 🚀 Deployment Preparation
```powershell
# Pre-deployment network validation
.\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,Intune,AzureAD -HtmlReport "Pre-Deployment-Check.html"
```

### 📊 Monitoring & Reporting
```powershell
# Scheduled task for regular reports
.\CheckMicrosoftEndpointsV2.ps1 -Services All -Quiet -HtmlReport "Weekly-Report-$(Get-Date -Format 'yyyy-MM-dd').html"

# Create performance baseline
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "Baseline-Performance.html"
```

---

## 🔧 Advanced Configuration

### ⚙️ Script Parameter Overview

#### CheckMicrosoftEndpointsV2.ps1
```powershell
# All available parameters
-Services        # All, WindowsUpdate, Intune, Defender, AzureAD, Microsoft365, Store, Activation, Edge, Telemetry, Interactive
-SkipPing        # Skip ping tests (faster)
-SkipSpeed       # Skip speed tests (faster)
-Quiet           # Silent mode (for automation)
-HtmlReport      # Path for HTML report
-OpenReport      # Automatically open report in browser
```

---

## 📈 Version History & Changelog

### 🏆 CheckMicrosoftEndpointsV2.ps1 Evolution

#### Version 2.1 (October 2025) - Current
- ✨ **HTML Report Generation** - Responsive design with Microsoft Look & Feel
- 🎯 **Service Selection** - Interactive menu + parameter-based selection
- ⚡ **Performance Options** - Configurable test depth
- 📱 **Mobile-Optimized** - Reports work on all devices
- 🔍 **Enhanced Analytics** - Detailed performance statistics

#### Version 2.0 (October 2025)
- 🎮 **Interactive Menu** - User-friendly service selection
- 📊 **Selective Testing** - Test only relevant services
- 🚀 **Speed Optimizations** - Skip optional tests
- 🔇 **Quiet Mode** - For automation and scripting

#### Version 1.1 (October 2025)
- 🏓 **Ping Latency Tests** - Network performance measurement
- 📈 **Download Speed Tests** - Bandwidth analysis
- 🎨 **Enhanced Output** - Color-coded results
- 📊 **Statistics** - Performance metrics and ratings

#### Version 1.0 (October 2025)
- 🔌 **Basic Connectivity** - TCP connection tests to Microsoft Services
- 🌐 **Service Coverage** - All major Microsoft Cloud Services
- 🎯 **Impact Analysis** - Effects of connection problems
- 📋 **Structured Output** - Organized results by services

---

---

## 🔒 Security & Permissions

### 🛡️ Required Permissions
| Script Category | Permissions | Reason |
|------------------|-------------|---------|
| **Network Tests** | Standard User | Test TCP connections only |
| **System Repair** | Administrator | Modify Windows components |
| **Intune/Graph** | Graph API | Manage cloud services |
| **Registry Operations** | Administrator | Modify system registry |

### 🔐 Best Practices
- **Least Privilege Principle** - Use minimal permissions only
- **Test Environment First** - Validate new scripts in test environment
- **Audit Logs** - Log important operations
- **Code Review** - Review scripts before production use

---

## 📚 Documentation & Help

### 📖 Script-specific Help
```powershell
# Show detailed help
Get-Help .\CheckMicrosoftEndpointsV2.ps1 -Full


# Parameter information
Get-Help .\CheckMicrosoftEndpointsV2.ps1 -Parameter Services
```


## 🤝 Contributing & Community

### 🌟 Contributing
1. **Fork** the repository
2. **Create feature branch** (`feature/amazing-feature`)
3. **Commit changes** (`git commit -m 'Add amazing feature'`)
4. **Push branch** (`git push origin feature/amazing-feature`)
5. **Create Pull Request**

### 📋 Contribution Guidelines
- **Code Style** - Follow PowerShell best practices
- **Documentation** - Update README and inline comments
- **Testing** - Validate scripts in test environment
- **Backwards Compatibility** - Consider compatibility with older versions

### 🐛 Bug Reports
Please use [GitHub Issues](https://github.com/roalhelm/PowershellScripts/issues) for:
- 🐛 Bug Reports
- 💡 Feature Requests  
- 📖 Documentation Improvements
- ❓ Questions and Discussions

---

## 📄 License & Credits

### 📜 License
This project is licensed under the [GNU General Public License v3.0](LICENSE) - see LICENSE file for details.

### 👨‍💻 Author
**Ronny Alhelm**
- 🌐 GitHub: [@roalhelm](https://github.com/roalhelm)
- 📧 Contact: Via GitHub Issues

### 🙏 Acknowledgments
- Microsoft Documentation & Best Practices
- PowerShell Community
- Enterprise IT Feedback
- Open Source Contributors

---

## 📊 Repository Stats

![Repository Stats](https://img.shields.io/badge/Scripts-25%2B-blue)
![PowerShell](https://img.shields.io/badge/Language-PowerShell-blue)
![Maintained](https://img.shields.io/badge/Maintained-Yes-green)
![Last Commit](https://img.shields.io/github/last-commit/roalhelm/PowershellScripts)

---

**💼 Ready for Enterprise Use** | **🚀 Continuously Updated** | **🛡️ Security Focused** | **📱 Modern Design**