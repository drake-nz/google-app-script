# Google Apps Script Projects

A multi-project Google Apps Script workspace for automating Google Sheets operations with rich text formatting preservation.

## 🎯 Overview

This repository contains Google Apps Script projects for:
- **Rich text concatenation** with formatting preservation
- **Sheet formatting** operations
- **Shared utilities** and configuration
- **Slides export** (planned)

## 📁 Project Structure

```
apps-script-projects/
├── formatting/           # Rich text concatenation & sheet formatting
│   ├── Code.gs          # Main concatenation logic
│   ├── appsscript.json  # Project configuration
│   └── *.md             # Documentation
├── shared-utilities/     # Common functions and config
│   ├── Config.gs        # Shared configuration
│   └── Utils.gs         # Utility functions
└── slides-export/        # Export to Google Slides (planned)
```

## ✨ Key Features

### Rich Text Concatenation
- Concatenate multiple cells while preserving individual character-level formatting
- Support for colors, font sizes, and styles from source cells
- Externalized configuration via Google Sheet (`Format_Config`)
- Handles cells with and without rich text formatting
- Automatic duplicate key resolution

### Configuration System
- Sheet-based configuration (no code changes needed)
- Flexible `ConfigKey` naming
- Process multiple concatenations in order
- Comprehensive error handling and logging

## 🚀 Quick Start

### Prerequisites
- [Node.js](https://nodejs.org/) installed
- [clasp](https://github.com/google/clasp) CLI tool:
  ```bash
  npm install -g @google/clasp
  ```

### Setup

1. **Clone this repository:**
   ```bash
   git clone <your-repo-url>
   cd apps-script-projects
   ```

2. **Login to Google:**
   ```bash
   clasp login
   ```

3. **Set up the formatting project:**
   ```bash
   cd formatting
   clasp create --title "Sheet Formatting" --type standalone
   clasp push
   ```

4. **Configure your Google Sheet:**
   - Create a `Format_Config` sheet in your Google Spreadsheet
   - Add configuration rows (see `formatting/CONCATENATION_CONFIG.md`)
   - Update `shared-utilities/Config.gs` with your Sheet ID

5. **Run concatenations:**
   - Open Apps Script editor: `clasp open-script`
   - Run `processAllConcatenations()` function

## 📖 Documentation

- **[Concatenation Configuration Guide](formatting/CONCATENATION_CONFIG.md)** - How to configure concatenations
- **[Setup Steps](formatting/SETUP_STEPS.md)** - Detailed setup instructions
- **[Troubleshooting](formatting/TROUBLESHOOTING.md)** - Common issues and solutions
- **[How to View Logs](formatting/HOW_TO_VIEW_LOGS.md)** - Debugging guide

## 🔧 Development Workflow

1. **Edit code** in your IDE
2. **Push changes:**
   ```bash
   cd formatting
   clasp push
   ```
3. **Test in browser:**
   ```bash
   clasp open-script
   ```
4. **View logs:**
   ```bash
   clasp logs
   ```

## 📝 Configuration Example

In your `Format_Config` sheet:

| ConfigKey | SourceSheet | SourceRanges | DestinationCell |
|-----------|-------------|--------------|-----------------|
| perfD8 | Performance and Progress | B8, C8 | D8 |
| siteG6 | Site Performance | E6, F6 | G6 |

Then run `processAllConcatenations()` to process all configurations.

## ⚠️ Important Notes

- **Never commit `.clasp.json`** - Contains personal authentication
- **Sheet IDs** - Store in `shared-utilities/Config.gs` or environment variables
- **Percentage symbols** - Google Sheets auto-detects `%` and may override rich text. Workaround: add a space before `%` (e.g., `35.1 %`)

## 🐛 Known Limitations

- Google Sheets auto-detects percentage values and may override rich text colors
- Some cells may not verify rich text immediately after setting (formatting still applies)
- Single-character cells with `%` may require special handling

## 📚 Technologies

- **Google Apps Script** - Server-side JavaScript for Google Workspace
- **clasp** - Command-line tool for Apps Script development
- **RichTextValueBuilder API** - For preserving character-level formatting

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

[Add your license here]

## 🙏 Acknowledgments

Built for automating Google Sheets formatting with rich text preservation.

---

**Note:** This repository is separate from other scripts and focuses specifically on Google Apps Script automation.
