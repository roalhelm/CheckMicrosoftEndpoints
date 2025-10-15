
# 🛠️ PowerShell Administrative Scripts Collection

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-GPL%20v3-green.svg)](LICENSE)
[![Last Update](https://img.shields.io/badge/Last%20Update-October%202025-brightgreen)](https://github.com/roalhelm/PowershellScripts)
[![Scripts](https://img.shields.io/badge/Scripts-25%2B-orange)](https://github.com/roalhelm/PowershellScripts)

Eine umfassende Sammlung von PowerShell-Skripten für Systemadministration, Microsoft Intune, Windows Updates, Benutzer-/Gruppenverwaltung, Netzwerk-Diagnostik und Remediation-Aufgaben in modernen Windows-Unternehmensumgebungen.

## 🌟 Highlights

- **🔌 Microsoft Endpoint Connectivity Tester V2.1** - Erweiterte Konnektivitäts-, Latenz- und Performance-Tests mit HTML-Reports

---

## 🚀 Quick Start

```powershell
# Repository klonen
git clone https://github.com/roalhelm/PowershellScripts.git
cd PowershellScripts

# Beispiel: Microsoft Endpoint Connectivity Test mit HTML-Report
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport -OpenReport

# Beispiel: Nur Intune und Defender testen
.\CheckMicrosoftEndpointsV2.ps1 -Services Intune,Defender -HtmlReport "MyReport.html"
```

---

## 📋 Systemanforderungen

| Anforderung | Details |
|-------------|---------|
| **PowerShell** | 5.1 oder höher |
| **Berechtigung** | Administratorrechte für die meisten Scripts |
| **Betriebssystem** | Windows 10/11, Windows Server 2016+ |
| **Module** | PSWindowsUpdate, Microsoft.Graph, ActiveDirectory |

---

## 🏆 Featured Script: Microsoft Endpoint Connectivity Tester V2.1

### ✨ Neue Funktionen in Version 2.1
- **🎨 HTML-Report Generation** - Professionelle, responsive Berichte
- **🎯 Service-Auswahl** - Teste nur relevante Microsoft Services
- **⚡ Performance-Optionen** - Überspringe Ping/Speed-Tests für schnellere Ausführung
- **📱 Responsive Design** - Reports funktionieren auf allen Geräten
- **🔍 Interaktives Menü** - Benutzerfreundliche Service-Auswahl

### 🎮 Verwendung

```powershell
# Interaktiver Modus (empfohlen für neue Benutzer)
.\CheckMicrosoftEndpointsV2.ps1

# Alle Services mit vollem Report
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport -OpenReport

# Schneller Test nur kritische Services
.\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,Intune,AzureAD -SkipSpeed

# Automatisierter Test für CI/CD
.\CheckMicrosoftEndpointsV2.ps1 -Services All -Quiet -HtmlReport "report.html"
```

### 🎯 Unterstützte Microsoft Services
- **Windows Update for Business** - Update-Endpunkte und Delivery-Services
- **Windows Autopatch** - Automatisches Patch-Management
- **Microsoft Intune** - Gerätemanagement und Compliance
- **Microsoft Defender** - Security und Threat Protection
- **Azure Active Directory** - Identitätsdienste und Authentication
- **Microsoft 365** - Office-Apps und Cloud-Services
- **Microsoft Store** - App-Verteilung und Updates
- **Windows Activation** - Lizenzierung und Aktivierung
- **Microsoft Edge** - Browser-Services und Enterprise-Features
- **Windows Telemetry** - Diagnosedaten und Fehlerberichterstattung

---

## 🎨 HTML-Report Features (CheckMicrosoftEndpointsV2.ps1)

### 📊 Dashboard-Übersicht
- **Statistische Karten** - Getestete Endpoints, Erfolgs-/Fehlerrate, Performance-Metriken
- **Farbkodierte Indikatoren** - Sofortige visuelle Bewertung
- **Responsive Grid-Layout** - Funktioniert auf Desktop, Tablet, Mobile

### 📋 Detaillierte Service-Tabellen
- **Service-spezifische Gruppierung** - Übersichtliche Organisation nach Microsoft Services
- **Status-Badges** - OK/FAILED mit Farbkodierung
- **IP-Adressen** - Für Netzwerk-Troubleshooting
- **Performance-Metriken** - Latenz und Geschwindigkeitsdaten (optional)

### 📈 Performance-Analyse
- **Latenz-Bewertung** - Automatische Einstufung (Exzellent/Gut/Verbesserungsbedarf)
- **Geschwindigkeits-Statistiken** - Min/Max/Durchschnitt
- **Service-Impact-Analyse** - Was bedeuten Ausfälle praktisch

### 🎨 Moderne Gestaltung
- **Microsoft Design Language** - Vertraute Optik für IT-Profis
- **Gradient-Header** - Professioneller Look
- **Schatten und Animationen** - Moderne Web-Ästhetik
- **Druckfreundlich** - Optimiert für PDF-Export

---

## 💼 Praxisbeispiele

### 🔧 Tägliche IT-Administration
```powershell
# Morgendlicher Netzwerk-Check mit Report
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "Daily-$(Get-Date -Format 'yyyy-MM-dd').html" -OpenReport

```

### 🚀 Deployment-Vorbereitung
```powershell
# Pre-Deployment Netzwerk-Validierung
.\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,Intune,AzureAD -HtmlReport "Pre-Deployment-Check.html"
```

### 📊 Monitoring & Reporting
```powershell
# Scheduled Task für regelmäßige Reports
.\CheckMicrosoftEndpointsV2.ps1 -Services All -Quiet -HtmlReport "Weekly-Report-$(Get-Date -Format 'yyyy-MM-dd').html"

# Performance-Baseline erstellen
.\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "Baseline-Performance.html"
```

---

## 🔧 Erweiterte Konfiguration

### ⚙️ Script-Parameter Übersicht

#### CheckMicrosoftEndpointsV2.ps1
```powershell
# Alle verfügbaren Parameter
-Services        # All, WindowsUpdate, Intune, Defender, AzureAD, Microsoft365, Store, Activation, Edge, Telemetry, Interactive
-SkipPing        # Ping-Tests überspringen (schneller)
-SkipSpeed       # Geschwindigkeitstests überspringen (schneller)
-Quiet           # Stiller Modus (für Automation)
-HtmlReport      # Pfad für HTML-Report
-OpenReport      # Report automatisch im Browser öffnen
```

---

## 📈 Version History & Changelog

### 🏆 CheckMicrosoftEndpointsV2.ps1 Evolution

#### Version 2.1 (Oktober 2025) - Current
- ✨ **HTML-Report Generation** - Responsive Design mit Microsoft Look & Feel
- 🎯 **Service-Auswahl** - Interactive Menu + Parameter-basierte Auswahl
- ⚡ **Performance-Optionen** - Konfigurierbare Test-Tiefe
- 📱 **Mobile-Optimiert** - Reports funktionieren auf allen Geräten
- 🔍 **Enhanced Analytics** - Detaillierte Performance-Statistiken

#### Version 2.0 (Oktober 2025)
- 🎮 **Interactive Menu** - Benutzerfreundliche Service-Auswahl
- 📊 **Selective Testing** - Teste nur relevante Services
- 🚀 **Speed Optimizations** - Überspringe optionale Tests
- 🔇 **Quiet Mode** - Für Automation und Scripting

#### Version 1.1 (Oktober 2025)
- 🏓 **Ping Latency Tests** - Netzwerk-Performance Messung
- 📈 **Download Speed Tests** - Bandbreiten-Analyse
- 🎨 **Enhanced Output** - Farbkodierte Ergebnisse
- 📊 **Statistics** - Performance-Metriken und Bewertungen

#### Version 1.0 (Oktober 2025)
- 🔌 **Basic Connectivity** - TCP-Verbindungstests zu Microsoft Services
- 🌐 **Service Coverage** - Alle wichtigen Microsoft Cloud Services
- 🎯 **Impact Analysis** - Auswirkungen von Verbindungsproblemen
- 📋 **Structured Output** - Organisierte Ergebnisse nach Services

---

---

## 🔒 Security & Permissions

### 🛡️ Erforderliche Berechtigungen
| Script-Kategorie | Berechtigungen | Begründung |
|------------------|----------------|------------|
| **Network Tests** | Standard User | Nur TCP-Verbindungen testen |
| **System Repair** | Administrator | Windows-Komponenten ändern |
| **Intune/Graph** | Graph API | Cloud-Services verwalten |
| **Registry Operations** | Administrator | System-Registry ändern |

### 🔐 Best Practices
- **Least Privilege Principle** - Nur minimale Berechtigungen verwenden
- **Test Environment First** - Neue Scripts in Testumgebung validieren
- **Audit Logs** - Wichtige Operationen protokollieren
- **Code Review** - Scripts vor Produktiv-Einsatz prüfen

---

## 📚 Dokumentation & Hilfe

### 📖 Script-spezifische Hilfe
```powershell
# Detaillierte Hilfe anzeigen
Get-Help .\CheckMicrosoftEndpointsV2.ps1 -Full


# Parameter-Informationen
Get-Help .\CheckMicrosoftEndpointsV2.ps1 -Parameter Services
```


## 🤝 Contributing & Community

### 🌟 Beitragen
1. **Fork** das Repository
2. **Feature Branch** erstellen (`feature/amazing-feature`)
3. **Änderungen committen** (`git commit -m 'Add amazing feature'`)
4. **Branch pushen** (`git push origin feature/amazing-feature`)
5. **Pull Request** erstellen

### 📋 Contribution Guidelines
- **Code Style** - PowerShell Best Practices befolgen
- **Documentation** - README und Inline-Comments aktualisieren
- **Testing** - Scripts in Testumgebung validieren
- **Backwards Compatibility** - Kompatibilität mit älteren Versionen beachten

### 🐛 Bug Reports
Bitte verwenden Sie die [GitHub Issues](https://github.com/roalhelm/PowershellScripts/issues) für:
- 🐛 Bug Reports
- 💡 Feature Requests  
- 📖 Dokumentations-Verbesserungen
- ❓ Fragen und Diskussionen

---

## 📄 License & Credits

### 📜 Lizenz
Dieses Projekt steht unter der [GNU General Public License v3.0](LICENSE) - siehe LICENSE-Datei für Details.

### 👨‍💻 Author
**Ronny Alhelm**
- 🌐 GitHub: [@roalhelm](https://github.com/roalhelm)
- 📧 Contact: Über GitHub Issues

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