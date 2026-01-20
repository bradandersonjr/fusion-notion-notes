# Fusion Notion Notes - Autodesk Fusion Add-in

A simple and efficient Autodesk Fusion add-in that provides quick access to create new Notion pages directly from the Autodesk Fusion interface. Perfect for designers and engineers who want to capture ideas, notes, or documentation without leaving their CAD environment.

## Features

- **One-click access** to create new Notion pages
- **Seamless integration** with Autodesk Fusion's Quick Access Toolbar
- **Cross-platform compatibility** (Windows & macOS)
- **Opens in default browser** for maximum compatibility
- **Clean error handling** with user-friendly messages
- **No additional dependencies** beyond Autodesk Fusion's built-in Python

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

1. **Click the "Fusion Notion Notes" button** in the Quick Access Toolbar (top of the Autodesk Fusion interface)
2. **A new Notion page will open** in your default web browser
3. **Start taking notes** or creating content immediately

## Technical Details

### Architecture
- Built using Autodesk Fusion's **Command and Event Handler** architecture
- Leverages Python's **webbrowser module** for cross-platform browser compatibility
- Follows Autodesk Fusion add-in **best practices** for UI integration
- Includes comprehensive **error handling** and cleanup procedures

### Key Components
- **NewNotionPageCommandExecuteHandler**: Handles the actual button click and opens Notion
- **NewNotionPageCommandCreatedHandler**: Sets up the command when Autodesk Fusion creates it
- **Utility functions**: Standardized error messaging and UI handling

### Code Structure
```
Fusion Notion Notes/
├── Fusion Notion Notes.py          # Main add-in code
├── Fusion Notion Notes.manifest    # Add-in configuration
├── resources/                    # UI resources
│   ├── 16x16-normal.png         # Small icon for the button
│   └── 32x32-normal.png         # Large icon for the button
└── README.md                    # This documentation
```

## Development

### Requirements
- **Autodesk Fusion** (any recent version with Python API support)
- **Python knowledge** for modifications (uses Python 3.7+ typically)
- **Basic understanding** of Autodesk Fusion's API architecture

### Customization
You can easily modify this add-in to:
- Open different Notion templates or pages
- Add custom shortcuts or hotkeys
- Integrate with other note-taking services
- Add additional UI elements or options

### Debugging
- Use Autodesk Fusion's **Text Commands** window to see error messages
- Check the **Scripts and Add-Ins** dialog for add-in status
- Review Python error messages in Autodesk Fusion's console

## Configuration

The add-in configuration is stored in `Fusion Notion Notes.manifest`:

```json
{
    "autodeskProduct": "Fusion360",
    "type": "addin",
    "id": "edae7804-57f3-433c-91b9-036230834e67",
    "author": "brad anderson jr",
    "description": {
        "": "Quick access to create new Notion pages from Autodesk Fusion"
    },
    "version": "1.0.0",
    "runOnStartup": false,
    "supportedOS": "windows|mac",
    "editEnabled": true
}
```

## Troubleshooting

### Common Issues

**Button doesn't appear in toolbar:**
- Ensure the add-in is activated in Scripts and Add-Ins dialog
- Restart Autodesk Fusion after installation
- Check that all files are in the correct directory

**Browser doesn't open:**
- Verify your default browser is set correctly
- Check internet connectivity
- Try running Autodesk Fusion as administrator (Windows)

**Error messages:**
- Review the error dialog for specific details
- Check Autodesk Fusion's Text Commands window for additional information
- Ensure Notion.so is accessible from your network

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
- GitHub: [@bradandersonjr](https://github.com/bradandersonjr)

## Version History

- **1.0.0** - Initial release with basic Notion page creation functionality
- **0.1.0** - Development version with core features

## Acknowledgments

- Built using Autodesk Fusion's Python API
- Inspired by the need for seamless note-taking in CAD workflows
- Thanks to the Autodesk Fusion developer community for API documentation and examples

---

*Made with ❤️ for the Autodesk Fusion community*
