# 3GPP Specifications Scraper

Automated scraper for 3GPP security specifications (33 series). Downloads, converts, and parses specifications into XML format.

## Features

- Downloads latest versions of 3GPP TS 33.xxx specifications
- Converts .doc files to .docx format using LibreOffice
- Parses specifications and extracts requirements/test cases into XML
- Automated updates via GitHub Actions (every 2 weeks)

## Local Usage

### Prerequisites

```bash
# Install LibreOffice
brew install --cask libreoffice  # macOS
# or
sudo apt-get install libreoffice  # Linux

# Install Python dependencies
pip install -r requirements.txt
```

### Run Scraper

```bash
python scraper.py
```

## GitHub Actions

The workflow runs automatically every 2 weeks and:
1. Downloads latest specifications
2. Converts documents to .docx and .xml
3. Commits changes if updates are found
