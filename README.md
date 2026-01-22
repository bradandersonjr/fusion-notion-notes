# Fusion Notion Notes - Autodesk Fusion Add-in

A professional Autodesk Fusion add-in that provides quick access to create new Notion pages directly from the Autodesk Fusion interface. Perfect for designers and engineers who want to capture ideas, notes, or documentation without leaving their CAD environment.

## Features

- **One-click QAT button** for instant Notion page creation
- **Configurable settings panel** with modern HTML-based interface
- **Dual-mode operation**: Open in web browser or Notion desktop app
- **Custom database URLs**: Direct links to your specific Notion databases
- **Automatic desktop app detection** with intelligent fallback to web browser
- **Persistent configuration** - settings save automatically and persist across sessions
- **No authentication required** - uses public Notion URLs
- **Cross-platform compatibility** (Windows & macOS)
- **Comprehensive error handling** with user-friendly messages

## Installation

1. **Download or clone** this repository to your local machine
2. **Copy the entire folder** to your Autodesk Fusion add-ins directory:
   - **Windows**: `%APPDATA%\Autodesk\Autodesk Fusion 360\API\AddIns\`
   - **macOS**: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`
3. **Start or restart** Autodesk Fusion
4. Open the **Scripts and Add-Ins** dialog:
   - Click the "Utilities" tab in the toolbar
   - Click "Scripts and Add-Ins" or press `Shift+S`
5. In the **Add-Ins** tab, find "Fusion Notion Notes"
6. Click **"Run"** to activate the add-in
7. The **Fusion Notion Notes** button will appear in your Quick Access Toolbar

## Usage

### Quick Start
1. **Click the "Fusion Notion Notes" button** in the Quick Access Toolbar
2. **Configure your settings** in the palette that appears:
   - Paste your Notion database URL
   - Choose between Web Browser or Desktop App
   - Click "Save Settings"
3. **Click the button again** to create a new page in your configured database

### Configuring Your Database URL

1. Open Notion in your browser or desktop app
2. Navigate to the database where you want to create new pages
3. Click the **3 dots (⋯)** in the top right corner of the database
4. Select **"Copy Link"** from the menu
5. Paste the URL into the settings panel in Fusion
6. Save your settings

### Settings Panel

Access the settings panel by:
- Clicking the "Fusion Notion Notes" button when no URL is configured
- Going to **Utilities → Scripts and Add-Ins → Fusion Notion Notes Settings**

## Technical Details

### Architecture
- Built on Autodesk Fusion's **Command and Event Handler** architecture
- Modern **HTML5-based settings interface** using Tailwind CSS
- **JSON-based configuration management** for persistent settings
- Cross-platform browser and desktop app integration via Python's `webbrowser` module
- **Protocol handler detection** for notion:// URL scheme
- Follows Autodesk Fusion add-in **best practices** for UI integration and cleanup

### Key Components

#### Python Backend
- **NotionQuickOpenHandler**: Handles QAT button clicks and opens Notion
- **NotionSettingsHandler**: Manages the settings palette lifecycle
- **PaletteCommandHandler**: Processes messages from HTML interface
- **Configuration System**: Load/save functions for persistent settings
- **Desktop App Detection**: Platform-specific protocol handler checking
- **Fallback System**: Automatic fallback to web browser when desktop app unavailable

#### HTML Frontend
- **Responsive settings interface** with modern design
- **Database URL configuration** with clear instructions
- **Open method selector** (Web Browser vs Desktop App)
- **Toast notifications** for user feedback
- **Icon system** powered by Lucide icons
- **Real-time configuration sync** between UI and backend

### Code Structure
```
Fusion Notion Notes/
├── Fusion Notion Notes.py          # Main add-in code (1200+ lines)
├── Fusion Notion Notes.manifest    # Add-in configuration
├── Palette.html                    # Settings UI interface
├── notion_config.json              # User configuration (auto-generated)
├── resources/                      # UI resources
│   ├── 16x16-normal.png           # Small icon for QAT button
│   └── 32x32-normal.png           # Large icon for QAT button
├── .gitignore                      # Git ignore rules
└── README.md                       # This documentation
```

### Configuration File Format

The `notion_config.json` file stores user preferences:
```json
{
  "database_url": "https://www.notion.so/your-workspace/database-id",
  "default_open_method": "web"
}
```

- `database_url`: The Notion database or page URL to open
- `default_open_method`: Either "web" or "desktop"

## Development

### Requirements
- **Autodesk Fusion** (any recent version with Python API support)
- **Python knowledge** for modifications (uses Python 3.7+)
- **Basic understanding** of Autodesk Fusion's API architecture
- **HTML/CSS/JavaScript** knowledge for UI modifications

### Customization

You can easily modify this add-in to:
- Change the palette UI design (edit `Palette.html`)
- Modify default settings (edit constants in Python file)
- Add additional configuration options
- Customize button icons (replace files in `resources/`)
- Add keyboard shortcuts
- Integrate with other services

### Debugging

- Use Autodesk Fusion's **Text Commands** window to see error messages
- Check the **Scripts and Add-Ins** dialog for add-in status
- Review Python error messages in the error dialog boxes
- Inspect `notion_config.json` for configuration issues
- Enable browser developer tools for HTML interface debugging

## Platform-Specific Features

### Windows
- Registry-based protocol handler detection for Notion desktop app
- Automatic fallback if desktop app not installed

### macOS
- Protocol handler support for Notion desktop app
- Seamless integration with macOS default browser

## Troubleshooting

### Common Issues

**Button doesn't appear in toolbar:**
- Ensure the add-in is activated in Scripts and Add-Ins dialog
- Restart Autodesk Fusion after installation
- Check that all files are in the correct add-ins directory

**Settings panel doesn't open:**
- Check that `Palette.html` exists in the add-in folder
- Review error messages in Autodesk Fusion's Text Commands window
- Try restarting the add-in

**Desktop app doesn't open:**
- Verify Notion desktop app is installed
- The add-in will automatically fall back to web browser
- You can manually select "Web Browser" in settings

**Configuration doesn't save:**
- Check file permissions in the add-in directory
- Ensure `notion_config.json` is not read-only
- Try manually deleting `notion_config.json` and reconfiguring

**Browser doesn't open:**
- Verify your default browser is set correctly in your OS
- Check internet connectivity
- Try running Autodesk Fusion as administrator (Windows)

## Contributing

1. **Fork** the repository
2. **Create a feature branch** (`git checkout -b feature/amazing-feature`)
3. **Commit your changes** (`git commit -m 'Add some amazing feature'`)
4. **Push to the branch** (`git push origin feature/amazing-feature`)
5. **Open a Pull Request**

## License

This project is open source and available under the [MIT License](https://opensource.org/licenses/MIT).

## Author

**Brad Anderson Jr**
- Email: brad@bradandersonjr.com
- Website: [bradandersonjr.com](https://www.bradandersonjr.com)
- Ko-fi: [ko-fi.com/bradandersonjr](https://ko-fi.com/bradandersonjr)

## Version History

- **0.6.0** - Current version
  - Modern HTML-based settings interface
  - Configurable database URLs
  - Desktop app vs web browser selection
  - Automatic desktop app detection with fallback
  - Persistent configuration storage
  - Top-docked settings palette
  - Toast notifications for user feedback

- **0.5.0** - Settings panel implementation
  - Added HTML palette for configuration
  - JSON-based config storage

- **0.1.0** - Initial development version
  - Basic QAT button with Notion page creation

## Acknowledgments

- Built using Autodesk Fusion's Python API
- UI powered by [Tailwind CSS](https://tailwindcss.com/)
- Icons provided by [Lucide](https://lucide.dev/)
- Inspired by the need for seamless note-taking in CAD workflows
- Thanks to the Autodesk Fusion developer community for API documentation and examples

---

*Made with ❤️ for the Autodesk Fusion community*
