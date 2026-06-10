# PC Repair DEX CRM V1 - Android App

Professional field service report generator for PC Repair DEX.

## Features
- Client management
- Automatic job numbering (DEX-S/M/E-YYYY-MM-XXXX)
- Small / Medium / Enterprise report tiers
- Dark/Light theme
- Photo support (before/after)
- Generates full .docx reports

## How to Build APK (Free on GitHub)

1. Push this repo to GitHub
2. The GitHub Action will automatically build the APK
3. Download from Actions tab → Artifacts

## Local Build
```bash
pip install buildozer
buildozer android debug
```

Place your template .docx files in the project if needed for advanced merging.

For questions, contact support.