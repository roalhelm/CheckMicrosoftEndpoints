# 📊 Microsoft Endpoint Connectivity Testing - Technische Dokumentation

## 📋 Übersicht

Dieses Dokument enthält umfassende technische Informationen über die Microsoft Endpoint Connectivity Tests, einschließlich der Datenquellen, Endpoint-URLs, Testmethodiken und technischen Implementierungsdetails.

---

## 🎯 Getestete Microsoft Services

### 📊 Service-Übersicht

| Service | Endpoints | Zweck | Kritikalität |
|---------|-----------|--------|--------------|
| **Windows Update for Business** | 16 URLs | System-Updates und Patches | 🔴 Kritisch |
| **Windows Autopatch** | 8 URLs | Automatische Patch-Verwaltung | 🟡 Hoch |
| **Microsoft Intune** | 16 URLs | Device Management/MDM | 🔴 Kritisch |
| **Microsoft Defender** | 11 URLs | Antivirus und Security | 🔴 Kritisch |
| **Azure Active Directory** | 8 URLs | Identity/Authentication | 🔴 Kritisch |
| **Microsoft 365** | 14 URLs | Produktivitäts-Suite | 🟡 Hoch |
| **Microsoft Store** | 8 URLs | App-Distribution | 🟢 Mittel |
| **Windows Activation** | 6 URLs | Lizenzierung | 🟡 Hoch |
| **Microsoft Edge** | 6 URLs | Browser-Services | 🟢 Mittel |
| **Windows Telemetry** | 10 URLs | Diagnostik und Monitoring | 🟢 Niedrig |

**Gesamt: 103 eindeutige Endpoints über 10 Microsoft Service-Kategorien**

---

## 🌐 Endpoint-Datenquellen und Methodik

### 📚 Offizielle Microsoft-Dokumentation

Die Endpoint-URLs werden aus folgenden **offiziellen Microsoft-Quellen** bezogen:

#### 🔗 Primäre Dokumentationsquellen

1. **Microsoft 365 Network Connectivity**
   - URL: https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges
   - Verwendung: Microsoft 365, Azure AD, Exchange Online

2. **Windows Update for Business**
   - URL: https://learn.microsoft.com/en-us/windows/deployment/update/waas-wu-settings
   - Verwendung: Windows Update, Delivery Optimization

3. **Microsoft Intune Network Requirements** 
   - URL: https://learn.microsoft.com/en-us/mem/intune/fundamentals/intune-endpoints
   - Verwendung: Intune Management, Enrollment, Compliance

4. **Microsoft Defender for Endpoint**
   - URL: https://learn.microsoft.com/en-us/microsoft-365/security/defender-endpoint/configure-proxy-internet
   - Verwendung: Defender Cloud Services, Threat Intelligence

5. **Azure Active Directory Connect**
   - URL: https://learn.microsoft.com/en-us/azure/active-directory/hybrid/reference-connect-ports
   - Verwendung: Azure AD, Device Registration, Authentication

6. **Microsoft Store for Business**
   - URL: https://learn.microsoft.com/en-us/microsoft-store/prerequisites-microsoft-store-for-business
   - Verwendung: Store Apps, Deployment, Licensing

7. **Windows Activation Services**
   - URL: https://learn.microsoft.com/en-us/windows/deployment/volume-activation/activate-using-key-management-service-vamt
   - Verwendung: KMS, MAK, Digital License Activation

#### 🔍 Zusätzliche Validierungsquellen

8. **Microsoft Edge Enterprise**
   - URL: https://learn.microsoft.com/en-us/deployedge/microsoft-edge-security-endpoints
   - Verwendung: Edge Updates, SmartScreen, Sync Services

9. **Windows Telemetry and Diagnostics**
   - URL: https://learn.microsoft.com/en-us/windows/privacy/configure-windows-diagnostic-data-in-your-organization
   - Verwendung: Diagnostic Data, Watson Error Reporting

10. **Windows Autopatch Documentation**
    - URL: https://learn.microsoft.com/en-us/windows/deployment/windows-autopatch/
    - Verwendung: Automated Patch Management, Device Health

### 🔄 Endpoint-Validierung und Aktualisierung

#### ✅ Validierungsmethodik

1. **Offizielle Dokumentation**: Alle URLs werden gegen Microsoft's offizielle Dokumentation validiert
2. **Network Testing**: Praktische Konnektivitätstests in Produktionsumgebungen
3. **Community Feedback**: Validierung durch Enterprise-IT-Communities
4. **Microsoft Support**: Bestätigung kritischer Endpoints durch Microsoft Support Cases

#### 🔄 Aktualisierungsprozess

- **Quarterly Reviews**: Vierteljährliche Überprüfung aller Endpoint-Listen
- **Microsoft Documentation Monitoring**: Überwachung von Änderungen in Microsoft Docs
- **Issue Tracking**: GitHub Issues für neue/veraltete Endpoints
- **Version Control**: Versionierung aller Endpoint-Änderungen

---

## 🏗️ Technische Implementierung

### 📊 Endpoint-Struktur im Script

```powershell
$backendUrls = @{
    'Windows Update for Business' = @(
        'https://dl.delivery.mp.microsoft.com',
        'https://download.windowsupdate.com',
        'https://fe3.delivery.mp.microsoft.com',
        # ... weitere URLs
    ) | Sort-Object -Unique
    
    'Intune' = @(
        'https://device.login.microsoftonline.com',
        'https://endpoint.microsoft.com',
        'https://graph.microsoft.com',
        # ... weitere URLs
    ) | Sort-Object -Unique
    
    # ... weitere Service-Kategorien
}
```

### 🔧 Testmethodiken

#### 1️⃣ **Konnektivitätstest (TCP)**
```powershell
Test-TcpConnectivity -HostName $hostname -Port 443
```
- **Zweck**: Socket-basierter TCP-Verbindungstest auf Port 443 (HTTPS)
- **Timeout**: 5 Sekunden pro Endpoint
- **Bewertung**: Erfolg/Fehler binär

#### 2️⃣ **Latenz-Test (Ping)**
```powershell
Test-Connection -ComputerName $hostname -Count 3 -Quiet
```
- **Zweck**: Netzwerk-Latenz und Packet-Loss Messung
- **Metriken**: Min/Max/Durchschnitt Latenz in ms
- **Bewertung**: Exzellent (<50ms), Gut (50-200ms), Verbesserungsbedarf (>200ms)

#### 3️⃣ **Performance-Test (Download)**
```powershell
Test-DownloadSpeed -Url $url
```
- **Zweck**: Response-Time und Server-Performance über WebClient und HEAD/GET-Fallback
- **Metriken**: HTTP Response Time in ms
- **Bewertung**: Schnell (<500ms), Normal (500-2000ms), Langsam (>2000ms)

### 🎨 HTML-Report Generation

#### 📊 Report-Struktur
- **Dashboard**: Übersicht mit Statistiken und Status-Karten
- **Service-Tabellen**: Detaillierte Ergebnisse gruppiert nach Microsoft Services
- **Performance-Charts**: Visuelle Latenz- und Speed-Analysen
- **Impact-Analysis**: Auswirkungen von Verbindungsproblemen

#### 🎨 Design-Elemente
- **Responsive Layout**: Bootstrap-ähnliches Grid-System
- **Microsoft Look & Feel**: Offizielle Microsoft-Farben und Fonts
- **Accessibility**: WCAG 2.1 konforme Farbkontraste
- **Mobile-Optimized**: Funktioniert auf allen Geräte-Größen

> Hinweis: Die V2-Konnektivitätstests laufen inzwischen auch unter PowerShell 7 auf macOS und Linux. Windows Activation wird dort automatisch übersprungen.

---

## 📋 Detaillierte Endpoint-Listen

### 🔄 Windows Update for Business (16 Endpoints)

| URL | Zweck | Kritikalität |
|-----|-------|--------------|
| `dl.delivery.mp.microsoft.com` | Content Delivery Network | Kritisch |
| `download.windowsupdate.com` | Update-Downloads | Kritisch |
| `fe3.delivery.mp.microsoft.com` | Frontend Delivery | Hoch |
| `sws.update.microsoft.com` | Software Update Services | Hoch |
| `update.microsoft.com` | Update-Katalog | Kritisch |
| `windowsupdate.microsoft.com` | Windows Update Client | Kritisch |
| `*.delivery.mp.microsoft.com` | Geo-distributed CDN | Hoch |
| `*.dl.delivery.mp.microsoft.com` | Download CDN | Hoch |
| `*.update.microsoft.com` | Regional Update Services | Hoch |
| `*.windowsupdate.com` | Legacy Update Services | Mittel |
| `*.windowsupdate.microsoft.com` | Regional WU Services | Hoch |

**Ausfallauswirkungen**: Keine Windows Updates, Sicherheitslücken, veraltete Treiber

### 🛡️ Microsoft Defender (11 Endpoints)

| URL | Zweck | Kritikalität |
|-----|-------|--------------|
| `*.defender.microsoft.com` | Defender Portal/Console | Hoch |
| `*.security.microsoft.com` | Security Center | Hoch |
| `*.wdcp.microsoft.com` | Windows Defender Cloud Protection | Kritisch |
| `definitionupdates.microsoft.com` | Virus Definition Updates | Kritisch |
| `unitedstates.cp.wd.microsoft.com` | US Cloud Protection | Kritisch |
| `europe.cp.wd.microsoft.com` | EU Cloud Protection | Kritisch |
| `asia.cp.wd.microsoft.com` | Asia Cloud Protection | Kritisch |

**Ausfallauswirkungen**: Veraltete Antivirus-Signaturen, kein Cloud-Schutz, eingeschränkte Threat Detection

### 📱 Microsoft Intune (13 Endpoints)

| URL | Zweck | Kritikalität |
|-----|-------|--------------|
| `device.login.microsoftonline.com` | Device Authentication | Kritisch |
| `endpoint.microsoft.com` | Endpoint Manager Portal | Hoch |
| `graph.microsoft.com` | Microsoft Graph API | Kritisch |
| `manage.microsoft.com` | Intune Management Service | Kritisch |
| `portal.manage.microsoft.com` | Intune Admin Portal | Hoch |
| `enrollment.manage.microsoft.com` | Device Enrollment | Kritisch |
| `enterpriseregistration.windows.net` | Azure AD Device Registration | Kritisch |

**Ausfallauswirkungen**: Kein Device Enrollment, App-Deployment schlägt fehl, Compliance-Checks funktionieren nicht

### 🔐 Azure Active Directory (8 Endpoints)

| URL | Zweck | Kritikalität |
|-----|-------|--------------|
| `login.microsoftonline.com` | Azure AD Authentication | Kritisch |
| `device.login.microsoftonline.com` | Device Login | Kritisch |
| `enterpriseregistration.windows.net` | Device Registration | Kritisch |
| `pas.windows.net` | Privileged Access Service | Hoch |
| `management.azure.com` | Azure Resource Manager | Hoch |
| `policykeyservice.dc.ad.msft.net` | Policy Key Service | Hoch |

**Ausfallauswirkungen**: Login-Probleme, Single Sign-On schlägt fehl, Device Registration nicht möglich

### 📊 Microsoft 365 (11 Endpoints)

| URL | Zweck | Kritikalität |
|-----|-------|--------------|
| `admin.microsoft.com` | Microsoft 365 Admin Center | Hoch |
| `config.office.com` | Office Configuration Service | Hoch |
| `graph.microsoft.com` | Microsoft Graph API | Kritisch |
| `officecdn.microsoft.com` | Office Content Delivery | Hoch |
| `protection.office.com` | Security & Compliance Center | Hoch |
| `portal.office.com` | Office 365 Portal | Mittel |
| `*.office.com` | Office Online Services | Hoch |
| `*.sharepoint.com` | SharePoint Online | Hoch |
| `*.onedrive.com` | OneDrive for Business | Hoch |

**Ausfallauswirkungen**: Office Apps funktionieren nicht vollständig, OneDrive Sync-Probleme, E-Mail-Zugriff eingeschränkt

---

## 🔧 Verwendung und Parameter

### 🎮 Interactive Mode (Standard)
```powershell
.\CheckMicrosoftEndpointsV2.ps1
```
Zeigt interaktives Menü zur Service-Auswahl

### 🚀 All Services mit vollem Report
```powershell
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "NetworkReport.html" -OpenReport
```

### ⚡ Schneller Test kritischer Services
```powershell
.\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,Intune,Defender,AzureAD -SkipSpeed
```

### 🤖 Automatisierter Test für CI/CD
```powershell
.\CheckMicrosoftEndpointsV2.ps1 -Services All -Quiet -HtmlReport "connectivity-report.html"
```

### 📊 Selektive Service-Tests
```powershell
# Nur Windows Update und Intune
.\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,Intune

# Nur Security-relevante Services  
.\CheckMicrosoftEndpointsV2.ps1 -Services Defender,AzureAD

# Nur Office/Produktivität
.\CheckMicrosoftEndpointsV2.ps1 -Services Microsoft365,Edge
```

---

## 🎨 HTML-Report Features

### 📊 Dashboard-Komponenten
- **Service Status Cards**: Übersicht aller getesteten Services
- **Performance Statistics**: Latenz-/Speed-Metriken
- **Failed Connections Summary**: Problembereiche sofort erkennbar
- **Test Configuration**: Welche Parameter wurden verwendet

### 📋 Detaillierte Service-Tabellen
- **Status Badges**: Farbkodierte OK/FAILED Indikatoren
- **IP Resolution**: Aufgelöste IP-Adressen für Troubleshooting
- **Performance Metrics**: Ping-Latenz und Response-Times (falls aktiviert)
- **Service Grouping**: Logische Gruppierung nach Microsoft Services

### 🎨 Design-Prinzipien
- **Microsoft Design Language**: Vertraute Optik für IT-Professionals
- **Responsive Layout**: Funktioniert auf Desktop, Tablet, Mobile
- **Accessibility**: Hohe Kontraste, Screen Reader kompatibel
- **Print-Friendly**: Optimiert für PDF-Export

---

## 🔍 Troubleshooting und Diagnostik

### ❌ Häufige Verbindungsprobleme

#### 🔥 Firewall/Proxy-Probleme
```
FAILED: Connection timed out
FAILED: Access denied
```
**Lösung**: Firewall-Regeln für HTTPS (443) zu Microsoft Domains prüfen

#### 🌐 DNS-Resolution-Probleme
```
FAILED: Name resolution failed
```
**Lösung**: DNS-Server-Konfiguration und Connectivity prüfen

#### ⚡ Performance-Probleme
```
OK but slow response (>2000ms)
High ping latency (>200ms)
```
**Lösung**: Netzwerk-Latenz und Bandbreite analysieren

### 🔧 Erweiterte Diagnostik

#### 📋 Log-Analyse
```powershell
# Detaillierte Fehlermeldungen aktivieren
$VerbosePreference = 'Continue'
.\CheckMicrosoftEndpointsV2.ps1 -Services All -Verbose
```

#### 🌐 Netzwerk-Trace
```powershell
# Netzwerk-Trace für spezifische URLs
netsh trace start capture=yes tracefile=network.etl provider=Microsoft-Windows-TCPIP
# Test ausführen
netsh trace stop
```

#### 📊 Performance-Baseline
```powershell
# Baseline-Report für Vergleiche erstellen
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "Baseline-$(Get-Date -Format 'yyyy-MM-dd').html"
```

---

## 📈 Performance-Metriken und Bewertung

### ⚡ Latenz-Bewertung (Ping)

| Kategorie | Latenz | Bewertung | Auswirkung |
|-----------|--------|-----------|------------|
| 🟢 **Exzellent** | < 50ms | Optimal | Keine Performance-Einbußen |
| 🟡 **Gut** | 50-200ms | Akzeptabel | Minimale Verzögerungen |
| 🔴 **Verbesserungsbedarf** | > 200ms | Problematisch | Spürbare Performance-Probleme |

### 🚀 Response-Time-Bewertung (HTTP)

| Kategorie | Response-Time | Bewertung | Auswirkung |
|-----------|---------------|-----------|------------|
| 🟢 **Schnell** | < 500ms | Optimal | Flüssige Nutzererfahrung |
| 🟡 **Normal** | 500-2000ms | Akzeptabel | Gelegentliche Verzögerungen |
| 🔴 **Langsam** | > 2000ms | Problematisch | Deutliche Performance-Einbußen |

### 📊 Service-spezifische SLAs

| Service | Acceptable Latenz | Critical Threshold | Business Impact |
|---------|------------------|-------------------|-----------------|
| **Windows Update** | < 100ms | > 500ms | Update-Downloads langsam |
| **Azure AD Login** | < 50ms | > 200ms | Login-Verzögerungen |
| **Intune Management** | < 150ms | > 1000ms | Device Management träge |
| **Defender Updates** | < 200ms | > 1000ms | Security-Risiko |

---

## 🔄 Automatisierung und Integration

### 📅 Scheduled Tasks
```powershell
# Täglicher Connectivity-Check
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\CheckMicrosoftEndpointsV2.ps1 -Services All -Quiet -HtmlReport 'C:\Reports\Daily-Connectivity.html'"
$trigger = New-ScheduledTaskTrigger -Daily -At "06:00"
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "Daily-Microsoft-Connectivity-Check"
```

### 🔄 CI/CD Integration
```yaml
# Azure DevOps Pipeline Step
- task: PowerShell@2
  displayName: 'Microsoft Endpoint Connectivity Test'
  inputs:
    targetType: 'filePath'
    filePath: '$(System.DefaultWorkingDirectory)/Scripts/CheckMicrosoftEndpointsV2.ps1'
    arguments: '-Services All -Quiet -HtmlReport "$(Agent.TempDirectory)/connectivity-report.html"'
```

### 📊 Monitoring Integration
```powershell
# SCOM/PRTG Integration
$exitCode = .\CheckMicrosoftEndpointsV2.ps1 -Services All -Quiet
if ($exitCode -ne 0) {
    # Alert senden
    Send-MonitoringAlert -Message "Microsoft Endpoint Connectivity Issues Detected"
}
```

---

## 🛡️ Security und Best Practices

### 🔒 Execution Policy
```powershell
# Sichere Execution Policy für Scripts
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 🔐 Least Privilege Principle
- **Standard User**: Basis-Connectivity-Tests
- **Administrator**: Nicht erforderlich für Network-Tests
- **Graph API**: Nur für erweiterte Intune-Diagnostik

### 🛡️ Firewall-Konfiguration
```powershell
# Windows Firewall Regel für ausgehende HTTPS-Verbindungen
New-NetFirewallRule -DisplayName "Microsoft Services HTTPS" -Direction Outbound -Protocol TCP -RemotePort 443 -Action Allow
```

---

## 📚 Weiterführende Ressourcen

### 📖 Offizielle Microsoft Dokumentation
- [Microsoft 365 Network Connectivity Principles](https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-network-connectivity-principles)
- [Intune Network Configuration](https://learn.microsoft.com/en-us/mem/intune/fundamentals/intune-endpoints)
- [Windows Update Delivery Optimization](https://learn.microsoft.com/en-us/windows/deployment/update/waas-delivery-optimization)
- [Azure AD Network Requirements](https://learn.microsoft.com/en-us/azure/active-directory/hybrid/reference-connect-ports)

### 🔧 Tools und Utilities
- [Microsoft Remote Connectivity Analyzer](https://testconnectivity.microsoft.com/)
- [Microsoft 365 Network Assessment Tool](https://connectivity.office.com/)
- [Windows Network Troubleshooter](https://support.microsoft.com/en-us/help/10741/windows-fix-network-connection-issues)

### 🏢 Enterprise Deployment Guides
- [Microsoft 365 Enterprise Deployment Guide](https://learn.microsoft.com/en-us/microsoft-365/enterprise/)
- [Intune Deployment Planning Guide](https://learn.microsoft.com/en-us/mem/intune/fundamentals/planning-guide)
- [Windows 10/11 Enterprise Deployment](https://learn.microsoft.com/en-us/windows/deployment/)

---

## 📄 Changelog und Versionierung

### Version 2.1 (Aktuell - Oktober 2025)
- ✨ **HTML-Report Generation** mit responsivem Design
- 🎯 **Service-Selection** Parameter und Interactive Menu  
- ⚡ **Performance-Optimierungen** mit Skip-Optionen
- 📱 **Mobile-Optimized** Reports
- 🔍 **Enhanced Analytics** und Performance-Statistiken

### Version 2.0 (Oktober 2025)  
- 🎮 **Interactive Menu-System** für Service-Auswahl
- 📊 **Selective Testing** Funktionalität
- 🚀 **Speed Optimizations** durch optionale Tests
- 🔇 **Quiet Mode** für Automation

### Version 1.1 (Oktober 2025)
- 🏓 **Ping Latency Tests** hinzugefügt
- 📈 **Download Speed Tests** implementiert  
- 🎨 **Enhanced Console Output** mit Farben
- 📊 **Performance Statistics** und Bewertungen

### Version 1.0 (Oktober 2025)
- 🔌 **Basic TCP Connectivity Tests** 
- 🌐 **Comprehensive Service Coverage**
- 🎯 **Impact Analysis** für failed connections
- 📋 **Structured Output** nach Service-Kategorien

---

## ⚠️ Microsoft Services Criticality Assessment

| Service | Endpoint Purpose | Criticality | Business Impact on Failure |
|---------|------------------|-------------|---------------------------|
| **Windows Update for Business** | Update Downloads, Catalog Sync, CDN Access | 🔴 **Critical** | No security updates, system vulnerabilities |
| **Microsoft Intune** | Device Management, Policy Sync, App Deployment | 🔴 **Critical** | No MDM, compliance loss, app installations fail |
| **Azure Active Directory** | Authentication, Token Renewal, Device Registration | 🔴 **Critical** | Login issues, SSO failure, device registration blocked |
| **Microsoft Defender** | Threat Intelligence, Cloud Protection, Definition Updates | 🟠 **High** | Outdated antivirus signatures, limited malware protection |
| **Windows Autopatch** | Automatic Patch Management, Update Orchestration | 🟠 **High** | Manual update management required, no automated deployments |
| **Microsoft 365** | Office Apps, OneDrive Sync, SharePoint Access | 🟠 **High** | Office apps limited, sync issues, productivity loss |
| **Windows Activation** | License Validation, KMS Services, Digital License | 🟡 **Medium** | Activation issues on fresh installs, license warnings |
| **Microsoft Store** | App Downloads, Updates, Deployment | 🟡 **Medium** | Store apps won't install, update issues with Store apps |
| **Microsoft Edge** | Browser Updates, SmartScreen, Enterprise Services | 🟡 **Medium** | Browser updates fail, limited security features |
| **Windows Telemetry** | Diagnostic Data, Error Reports, Analytics | 🟢 **Low** | No functional impact, only diagnostic data affected |

### 🎯 Criticality Legend

| Symbol | Level | Downtime Tolerance | Action Required |
|--------|-------|-------------------|-----------------|
| 🔴 **Critical** | System-essential | < 1 Hour | Immediate fix required |
| 🟠 **High** | Business-critical | < 4 Hours | Priority treatment |
| 🟡 **Medium** | Functionality-impact | < 24 Hours | Planned remediation |
| 🟢 **Low** | Optional/Analytics | > 24 Hours | Fix when convenient |

---

## 👨‍💻 Author & Maintenance

**Entwickelt von**: Ronny Alhelm  
**GitHub**: [@roalhelm](https://github.com/roalhelm)  
**Repository**: [PowershellScripts](https://github.com/roalhelm/PowershellScripts)  
**Lizenz**: GNU General Public License v3.0  

**Letzte Aktualisierung**: Oktober 2025  
**Nächste Review**: Januar 2026  

---
