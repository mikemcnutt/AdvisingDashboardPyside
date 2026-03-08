# Advising Dashboard

A beautiful glassmorphism-styled desktop application for managing student advising workflows.

## Features

- 🎨 Modern glassmorphism UI design
- 📊 Track student advising status across semesters
- 📧 Email integration with Outlook
- 🔍 Search and filter students by track
- 💾 Auto-save settings and window state
- 🖥️ Optimized for 1920x1080 displays

## Download

### Pre-built Windows Executable

1. Go to the [Releases](../../releases) page
2. Download `AdvisingDashboard.exe` from the latest release
3. Run the executable - no installation required!

### Build from Source

**Requirements:**
- Python 3.11+
- PySide6
- pywin32 (for Outlook integration on Windows)

**Installation:**

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/advising-dashboard.git
cd advising-dashboard

# Install dependencies
pip install -r requirements.txt

# Run the application
python advising_dashboard_glass.py
```

## Building Your Own Executable

### Using GitHub Actions (Automatic)

1. Push your code to GitHub
2. GitHub Actions will automatically build the .exe file
3. Download from the "Actions" tab under "Artifacts"

### Local Build with PyInstaller

```bash
# Install PyInstaller
pip install pyinstaller

# Build the executable
pyinstaller --name="AdvisingDashboard" --onefile --windowed advising_dashboard_glass.py

# Find the .exe in the dist/ folder
```

## Usage

1. **Launch**: Run `AdvisingDashboard.exe`
2. **Set Folder**: Click "Browse" and select your advising folder containing student JSON files
3. **Select Semesters**: Check Spring, Summer, and/or Fall
4. **Scan**: Click "Scan Folder" to load students
5. **Manage**: Students automatically categorize into three columns:
   - **Needs Advised**: Students requiring appointments
   - **Advised (Not Complete)**: Partially advised students
   - **Advised**: Fully advised students

## File Requirements

The dashboard expects student data in JSON format with the following structure:

```json
{
  "studentName": "John Doe",
  "studentID": "123456789",
  "kctcsEmail": "john.doe@kctcs.edu",
  "personalEmail": "john@example.com",
  "track": "NA",
  "notes": "Optional notes",
  "semesterPlan": {
    "spring": {
      "courses": [...],
      "declined": false,
      "notComplete": false
    },
    "summer": {...},
    "fall": {...}
  }
}
```

## Configuration

Settings are automatically saved to `advising_dashboard_settings.json` including:
- Last used folder path
- Selected year and semesters
- Email subject and scheduling link
- Window size and position
- Track filter preference

## Email Integration

**Requirements**: Microsoft Outlook installed on Windows

**Features**:
- Bulk email to students needing advising
- Create drafts for review before sending
- Customizable subject line and scheduling link
- Individual email buttons for partially advised students

## System Requirements

- **OS**: Windows 10/11 (for .exe version)
- **Resolution**: 1280x720 minimum, optimized for 1920x1080
- **RAM**: 4GB minimum
- **Disk**: 100MB free space

## Troubleshooting

**"Could not start local server"**
- The HTML advising interface file must be in the same folder as the executable
- File should be named `advising.html` or `Advising.html`

**Email features not working**
- Install Microsoft Outlook
- Ensure Outlook is set as default email client

**Window too small/large**
- Delete `advising_dashboard_settings.json` to reset window size
- Or manually maximize/resize and settings will be saved

## License

[Your chosen license here]

## Contributing

Pull requests welcome! Please ensure:
- Code follows PEP 8 style guidelines
- Test on Windows before submitting
- Update README if adding features

## Credits

Built with:
- **PySide6** - Qt for Python UI framework
- **PyInstaller** - Executable packaging
- **pywin32** - Windows COM automation for Outlook
