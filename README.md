<!--MODERNIZED:v2-->
# Quick Word Recovr

> Cross-platform DOCX recovery. Originally a Windows-only tool migrated from SourceForge; now also available as a fully offline web/PWA app for every major platform.

[![Live app](https://img.shields.io/badge/live-app-ff2e93?style=for-the-badge)](https://socrtwo.github.io/quickwordrecovr-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/quickwordrecovr-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/quickwordrecovr-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/quickwordrecovr-SF?style=for-the-badge&color=22d3ee)](https://github.com/socrtwo/quickwordrecovr-SF/blob/main/LICENSE)
[![Last commit](https://img.shields.io/github/last-commit/socrtwo/quickwordrecovr-SF?style=for-the-badge&color=34d399)](https://github.com/socrtwo/quickwordrecovr-SF/commits)

**Live app:** https://socrtwo.github.io/quickwordrecovr-SF/
**Downloads:** [Releases](https://github.com/socrtwo/quickwordrecovr-SF/releases)
**Source:** [socrtwo/quickwordrecovr-SF](https://github.com/socrtwo/quickwordrecovr-SF)

---

## What it does

When Microsoft Word reports *"unspecified error"* opening a `.docx`, the file's
internal `word/document.xml` is usually truncated or malformed. Quick Word Recovr:

1. Treats the `.docx` as a ZIP archive and reads `word/document.xml`.
2. Validates the XML and locates the last well-formed paragraph.
3. Rebuilds a clean document with proper closing tags.
4. Repacks the archive — or extracts plain text as a fallback.

The web app does all of this **locally in your browser**. Files are never uploaded.

## Platforms

| Platform                | Format                            | How to install                                              |
| ----------------------- | --------------------------------- | ----------------------------------------------------------- |
| **Web**                 | hosted PWA                        | Open the [live app](https://socrtwo.github.io/quickwordrecovr-SF/) |
| **Windows**             | `.exe` + Inno Setup installer     | Download from [Releases](../../releases)                    |
| **macOS**               | PWA bundle (`.zip`)               | Extract, open `index.html`, or install via browser          |
| **Linux**               | PWA bundle (`.zip`)               | Extract and serve, or install as PWA from the live site     |
| **ChromeOS**            | PWA                               | Open the live app, then *Install* from the address bar      |
| **Android**             | PWA bundle (`.zip`) / installable | Open the live app in Chrome, then *Add to Home screen*      |
| **iOS / iPadOS**        | PWA bundle (`.zip`) / installable | Open the live app in Safari, then *Share → Add to Home Screen* |

The web app works **fully offline** once installed — recovery runs in JavaScript
using [JSZip](https://stuk.github.io/jszip/) and the browser's native XML parser.

## Repository layout

```
.
├── Unspecified Error DOCX Recovery/   VB.NET WinForms source (.NET Framework 4.8)
├── web/                                Cross-platform PWA (JS/HTML/CSS)
│   ├── index.html                      DOCX recovery tool
│   ├── about.html                      README-driven landing
│   ├── app.js, app.css                 App logic + styles
│   ├── manifest.webmanifest, sw.js     PWA manifest + offline service worker
│   └── icons/                          App icons (SVG + PNG)
├── setupVII-without-adware-offers.iss  Inno Setup installer script
├── releases/                           Legacy SourceForge archives
└── .github/workflows/                  Build / Pages / Release CI
```

## Building

### Windows (VB.NET)

Requires Visual Studio 2019+ or MSBuild with .NET Framework 4.8 SDK.

```powershell
nuget restore "Unspecified Error DOCX Recovery.sln"
msbuild "Unspecified Error DOCX Recovery.sln" /p:Configuration=Release
```

The CI workflow (`.github/workflows/build.yml`) builds this on every push.

### Web / PWA

No build step. Serve the `web/` directory with any static HTTP server:

```bash
cd web && python3 -m http.server 8080
```

Then open <http://localhost:8080/>.

### Installer (Inno Setup)

`setupVII-without-adware-offers.iss` produces a Windows installer.

## Origin

Originally hosted on SourceForge and migrated to GitHub.

- **SourceForge:** https://sourceforge.net/projects/quickwordrecovr/
- **Migration tool:** [SF2GH Migrator](https://github.com/socrtwo/sf-to-github)

## Contributing

Issues and pull requests welcome at
<https://github.com/socrtwo/quickwordrecovr-SF/issues>.

## License

MIT — see [LICENSE](LICENSE).

---

*Maintained by [@socrtwo](https://github.com/socrtwo)*
