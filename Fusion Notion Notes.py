"""
Fusion Notion Notes - Autodesk Fusion Add-in
=====================================

A simple Autodesk Fusion add-in that provides quick access to create new Notion pages
directly from the Autodesk Fusion interface. This add-in adds a button to the Quick Access
Toolbar (QAT) that opens a beautiful HTML palette with options for web or desktop.

Author: Brad Anderson Jr
Contact: brad@bradandersonjr.com
Version: 0.5.0

Features:
- One-click QAT button to create new Notion pages
- Settings panel in Scripts menu to configure database and open method
- Choose default behavior: web browser or desktop app
- Configuration saved locally (no login or API authentication required)
- Clean error handling with user-friendly messages

Installation:
1. Copy this add-in folder to your Autodesk Fusion add-ins directory
2. Start/restart Autodesk Fusion
3. Go to Scripts and Add-Ins dialog
4. Select this add-in and click "Run"
5. The Fusion Notion Notes button will appear in your Quick Access Toolbar

Usage:
- Click "Fusion Notion Notes" button to open the palette
- Choose to open in Web Browser or Desktop App
- Configure your database URL in settings
- If configured, new pages will open in your specified database
- If not configured, new pages will open in your default Notion location

Configuration:
- The database URL is saved in a local JSON file (notion_config.json)
- No authentication or API keys required
- Simply paste the URL of your Notion database in the settings dialog

Technical Details:
- Uses Autodesk Fusion's Command and Event Handler architecture
- Custom HTML palette for beautiful UI
- Leverages Python's webbrowser module for cross-platform compatibility
- Supports both HTTPS (web) and notion:// (desktop app) protocols
- Configuration stored in JSON format in the add-in directory
- Follows Autodesk Fusion add-in best practices for UI integration
- Includes comprehensive error handling and cleanup procedures
"""

import adsk.core, adsk.fusion, traceback, webbrowser, json, os, subprocess, platform

# ============================================================================
# CONSTANTS AND CONFIGURATION
# ============================================================================
ADDIN_NAME = 'Fusion Notion Notes'  # Display name for the add-in
ADDIN_VERSION = '0.6.0'  # Current version number
ADDIN_AUTHOR = 'Brad Anderson Jr'  # Add-in author
ADDIN_CONTACT = 'brad@bradandersonjr.com'  # Contact information
ERROR_MSG = 'Failed:\n{}'  # Template for error message formatting
CONFIG_FILENAME = 'notion_config.json'  # Configuration file name
PALETTE_ID = 'FusionNotionNotesPalette'  # Unique ID for the palette
SETTINGS_CMD_ID = 'FusionNotionNotesSettings'  # Command ID for settings

# ============================================================================
# GLOBAL VARIABLES
# ============================================================================

# List to hold event handlers to prevent them from being garbage collected
handlers = []

# Reference to the palette
palette = None

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def get_config_path():
    """Gets the full path to the configuration file."""
    addin_dir = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(addin_dir, CONFIG_FILENAME)


def load_config():
    """Loads the configuration from the JSON file."""
    config_path = get_config_path()
    if not os.path.exists(config_path):
        # Create default config file
        default_config = {
            "database_url": "https://www.notion.so/new",
            "default_open_method": "web"
        }
        save_config(default_config)
        return default_config
    try:
        with open(config_path, 'r') as f:
            return json.load(f)
    except Exception:
        # If loading fails, return default config
        default_config = {
            "database_url": "https://www.notion.so/new",
            "default_open_method": "web"
        }
        return default_config


def save_config(config):
    """Saves the configuration to the JSON file."""
    config_path = get_config_path()
    try:
        with open(config_path, 'w') as f:
            json.dump(config, f, indent=2)
        return True
    except Exception:
        return False


def get_notion_url(protocol='https'):
    """Gets the appropriate Notion URL based on configuration."""
    config = load_config()
    database_url = config.get('database_url', '').strip()

    if database_url:
        # User has configured a database URL
        if protocol == 'notion':
            # Convert https:// to notion://
            if database_url.startswith('https://'):
                return database_url.replace('https://', 'notion://', 1)
            elif database_url.startswith('http://'):
                return database_url.replace('http://', 'notion://', 1)
        return database_url
    else:
        # No database configured, use default new page URL
        if protocol == 'notion':
            return 'notion://www.notion.so/new'
        else:
            return 'https://www.notion.so/new'


def show_error_message(ui, error_message):
    """Displays an error message dialog to the user."""
    if ui:
        ui.messageBox(error_message, ADDIN_NAME, 0, 0)


def check_notion_protocol_handler():
    """Checks if the Notion protocol handler (notion://) is available.
    
    Returns:
        True if protocol handler exists, False otherwise
    """
    try:
        if platform.system() == 'Windows':
            # On Windows, check registry for notion:// protocol handler
            try:
                # Use reg query to check if notion protocol is registered
                result = subprocess.run(
                    ['reg', 'query', 'HKEY_CLASSES_ROOT\\notion', '/ve'],
                    capture_output=True,
                    timeout=2,
                    creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                )
                return result.returncode == 0
            except Exception:
                # If registry check fails, assume protocol might not exist
                return False
        else:
            # On macOS/Linux, we can't easily check, so assume it might exist
            # The webbrowser module will handle it
            return True
    except Exception:
        return False


def open_notion_with_fallback(protocol='https', ui=None):
    """Opens Notion URL with fallback to web if desktop app is not available.
    
    Args:
        protocol: 'notion' for desktop app, 'https' for web browser
        ui: User interface object for showing messages (optional)
    
    Returns:
        True if opened successfully, False otherwise
    """
    try:
        url = get_notion_url(protocol=protocol)
        
        if protocol == 'notion':
            # Check if protocol handler exists before trying
            if not check_notion_protocol_handler():
                # Protocol handler doesn't exist, fall back to web browser
                web_url = get_notion_url(protocol='https')
                webbrowser.open_new(web_url)
                if ui:
                    ui.messageBox(
                        'Notion desktop app not found. Opened in web browser instead.',
                        ADDIN_NAME,
                        0,  # OK button only
                        0   # Information icon
                    )
                return False
            
            # Try to open with desktop app
            try:
                result = webbrowser.open(url)
                
                # On some systems, webbrowser.open() might return False on failure
                # But on Windows, it often returns True even if protocol handler doesn't work well
                # So we rely on the registry check above
                return True
                
            except Exception:
                # If opening fails, fall back to web browser
                web_url = get_notion_url(protocol='https')
                webbrowser.open_new(web_url)
                if ui:
                    ui.messageBox(
                        'Could not open Notion desktop app. Opened in web browser instead.',
                        ADDIN_NAME,
                        0,  # OK button only
                        0   # Information icon
                    )
                return False
        else:
            # Open in web browser
            webbrowser.open_new(url)
            return True
            
    except Exception as e:
        # If everything fails, try web browser as last resort
        try:
            web_url = get_notion_url(protocol='https')
            webbrowser.open_new(web_url)
            if ui and protocol == 'notion':
                ui.messageBox(
                    'Could not open Notion desktop app. Opened in web browser instead.',
                    ADDIN_NAME,
                    0,
                    0
                )
        except Exception:
            pass  # If even web browser fails, we're out of options
        return False


# ============================================================================
# PALETTE HANDLER
# ============================================================================

class PaletteCommandHandler(adsk.core.HTMLEventHandler):
    """Handles events from the HTML palette."""

    def __init__(self, ui):
        super().__init__()
        self.ui = ui

    def notify(self, args):
        """Called when the HTML palette sends a message."""
        try:
            htmlArgs = adsk.core.HTMLEventArgs.cast(args)
            action = htmlArgs.action

            if action == 'getConfig':
                # Send current configuration to the palette
                config = load_config()
                database_url = config.get('database_url', '')
                default_method = config.get('default_open_method', 'web')
                return_data = json.dumps({
                    'action': 'setConfig',
                    'databaseUrl': database_url,
                    'defaultMethod': default_method
                })
                htmlArgs.returnData = return_data
                # Also proactively send via sendInfoToHTML as backup
                try:
                    global palette
                    if palette and palette.isVisible:
                        palette.sendInfoToHTML('setConfig', return_data)
                except Exception:
                    pass  # Silently fail if palette not ready

            elif action == 'savePreferences':
                # Save preferences from the settings panel
                data = json.loads(htmlArgs.data) if htmlArgs.data else {}
                config = load_config()

                database_url = data.get('databaseUrl', '').strip()
                default_method = data.get('defaultMethod', 'web')

                config['database_url'] = database_url
                config['default_open_method'] = default_method

                save_config(config)

            elif action == 'openNotionForUrl':
                # Open Notion in web browser so user can navigate and get database URL
                webbrowser.open_new('https://www.notion.so')
            
            elif action == 'openUrl':
                # Open a URL in the user's browser
                url = htmlArgs.data if htmlArgs.data else ''
                if url:
                    webbrowser.open_new(url)

        except Exception as e:
            show_error_message(self.ui, ERROR_MSG.format(traceback.format_exc()))


# ============================================================================
# COMMAND HANDLERS
# ============================================================================

class NotionQuickOpenHandler(adsk.core.CommandEventHandler):
    """Handles opening Notion directly from the QAT button."""

    def __init__(self, ui):
        super().__init__()
        self.ui = ui

    def notify(self, args):
        """Called when the user clicks the Fusion Notion Notes button."""
        try:
            config = load_config()
            default_method = config.get('default_open_method', 'web')

            if default_method == 'desktop':
                # Try desktop app with fallback to web
                open_notion_with_fallback(protocol='notion', ui=self.ui)
            else:
                # Open in web browser
                open_notion_with_fallback(protocol='https', ui=self.ui)

        except Exception as e:
            show_error_message(self.ui, ERROR_MSG.format(traceback.format_exc()))


def send_config_to_palette(palette_instance):
    """Helper function to send current config to the palette."""
    if palette_instance and palette_instance.isVisible:
        try:
            config = load_config()
            database_url = config.get('database_url', '')
            default_method = config.get('default_open_method', 'web')
            config_data = json.dumps({
                'action': 'setConfig',
                'databaseUrl': database_url,
                'defaultMethod': default_method
            })
            palette_instance.sendInfoToHTML('setConfig', config_data)
        except Exception:
            pass  # Silently fail if palette is not ready


class PaletteClosedHandler(adsk.core.UserInterfaceGeneralEventHandler):
    """Handles palette closed events - when reopened, config will be sent."""
    
    def __init__(self, ui):
        super().__init__()
        self.ui = ui
    
    def notify(self, args):
        """Called when palette is closed."""
        try:
            # When palette is closed, we don't need to do anything
            # Config will be sent when it's reopened via NotionSettingsHandler
            pass
        except Exception:
            pass  # Silently fail


class NotionSettingsHandler(adsk.core.CommandEventHandler):
    """Handles showing the settings palette."""

    def __init__(self, ui):
        super().__init__()
        self.ui = ui

    def notify(self, args):
        """Called when the user clicks the settings command."""
        try:
            global palette

            if palette:
                # Toggle palette visibility
                was_visible = palette.isVisible
                palette.isVisible = not was_visible
                # If palette is now visible, send config
                if palette.isVisible:
                    send_config_to_palette(palette)
            else:
                # Create the palette if it doesn't exist
                palette = self.ui.palettes.itemById(PALETTE_ID)
                if not palette:
                    # Get the HTML file path
                    addin_dir = os.path.dirname(os.path.realpath(__file__))
                    html_file = os.path.join(addin_dir, 'palette.html')

                    # Convert Windows path to file:// URL format
                    # Replace backslashes with forward slashes for file:// URL
                    html_file_url = html_file.replace('\\', '/')

                    # Create the palette
                    palette = self.ui.palettes.add(
                        PALETTE_ID,
                        'Fusion Notion Notes Settings',
                        html_file_url,
                        True,  # Show palette
                        True,  # Show close button
                        True,  # Can be resized
                        720,   # Width
                        900    # Height
                    )

                    # Add HTML event handler
                    onHTML = PaletteCommandHandler(self.ui)
                    palette.incomingFromHTML.add(onHTML)
                    handlers.append(onHTML)

                    # Add closed event handler to detect visibility changes
                    onClosed = PaletteClosedHandler(self.ui)
                    palette.closed.add(onClosed)
                    handlers.append(onClosed)

                    # Dock the palette to the right side
                    palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateFloating

                    # Send initial config to palette immediately and after delays to ensure HTML is ready
                    send_config_to_palette(palette)
                    # Also send after short delays to ensure HTML has loaded
                    def send_config_delayed():
                        send_config_to_palette(palette)
                    # Use a timer to send config after HTML loads (Autodesk Fusion doesn't have threading, so we'll rely on multiple sends)
                    # The HTML will handle duplicate configs gracefully
                else:
                    palette.isVisible = True
                    # Send config when showing existing palette
                    send_config_to_palette(palette)

        except Exception as e:
            show_error_message(self.ui, ERROR_MSG.format(traceback.format_exc()))


class NotionQuickOpenCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """Handles the creation of the Notion Quick Open command (QAT button)."""

    def __init__(self, ui):
        super().__init__()
        self.ui = ui

    def notify(self, args):
        """Called when the command is created."""
        try:
            command = args.command

            # Add execute handler
            onExecute = NotionQuickOpenHandler(self.ui)
            command.execute.add(onExecute)
            handlers.append(onExecute)

        except Exception as e:
            show_error_message(self.ui, ERROR_MSG.format(traceback.format_exc()))


class NotionSettingsCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """Handles the creation of the Notion Settings command."""

    def __init__(self, ui):
        super().__init__()
        self.ui = ui

    def notify(self, args):
        """Called when the command is created."""
        try:
            command = args.command

            # Add execute handler
            onExecute = NotionSettingsHandler(self.ui)
            command.execute.add(onExecute)
            handlers.append(onExecute)

        except Exception as e:
            show_error_message(self.ui, ERROR_MSG.format(traceback.format_exc()))


# ============================================================================
# MAIN ADD-IN FUNCTIONS
# ============================================================================

def run(context):
    """Entry point for the Autodesk Fusion add-in."""
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        # Get reference to the Quick Access Toolbar
        qatToolbar = ui.toolbars.itemById('QAT')

        # Create QAT command (quick open button)
        qat_command_id = f'{ADDIN_NAME}CmdDef'
        notionQuickOpenCmdDef = ui.commandDefinitions.addButtonDefinition(
            qat_command_id,
            ADDIN_NAME,
            'Create a new Notion page',
            './resources'
        )

        # Create and attach command created handler for QAT button
        onQuickOpenCreated = NotionQuickOpenCommandCreatedHandler(ui)
        notionQuickOpenCmdDef.commandCreated.add(onQuickOpenCreated)
        handlers.append(onQuickOpenCreated)

        # Add the quick open button to the Quick Access Toolbar
        qatToolbar.controls.addCommand(notionQuickOpenCmdDef, 'HealthStatusCommand', False)

        # Create Settings command for Scripts menu
        notionSettingsCmdDef = ui.commandDefinitions.addButtonDefinition(
            SETTINGS_CMD_ID,
            'Fusion Notion Notes Settings',
            'Configure Notion database and default open method',
            './resources'
        )

        # Create and attach command created handler for settings
        onSettingsCreated = NotionSettingsCommandCreatedHandler(ui)
        notionSettingsCmdDef.commandCreated.add(onSettingsCreated)
        handlers.append(onSettingsCreated)

        # Add settings to the ADD-INS panel in UTILITIES toolbar
        workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        if workspace:
            addInsPanel = workspace.toolbarPanels.itemById('SolidScriptsAddinsPanel')
            if addInsPanel:
                control = addInsPanel.controls.addCommand(notionSettingsCmdDef)
                # Check if control is a dropdown and disable it
                if control:
                    try:
                        # Try to access dropdown-specific properties
                        if hasattr(control, 'isDropDown'):
                            control.isDropDown = False
                    except:
                        pass

    except Exception as e:
        if ui:
            show_error_message(ui, ERROR_MSG.format(traceback.format_exc()))


def stop(context):
    """Cleanup function called when the add-in is stopped."""
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        # Hide and delete the palette
        global palette
        if palette:
            palette.deleteMe()
            palette = None

        # Get reference to the Quick Access Toolbar
        qatToolbar = ui.toolbars.itemById('QAT')

        # Delete QAT command
        qat_command_id = f'{ADDIN_NAME}CmdDef'
        cmdDef = ui.commandDefinitions.itemById(qat_command_id)
        if cmdDef:
            cmdDef.deleteMe()

        cmd = qatToolbar.controls.itemById(qat_command_id)
        if cmd:
            cmd.deleteMe()

        # Delete Settings command
        settingsCmdDef = ui.commandDefinitions.itemById(SETTINGS_CMD_ID)
        if settingsCmdDef:
            settingsCmdDef.deleteMe()

        # Remove from ADD-INS panel
        workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        if workspace:
            addInsPanel = workspace.toolbarPanels.itemById('SolidScriptsAddinsPanel')
            if addInsPanel:
                settingsControl = addInsPanel.controls.itemById(SETTINGS_CMD_ID)
                if settingsControl:
                    settingsControl.deleteMe()

        # Clear handlers
        handlers.clear()

    except:
        if ui:
            show_error_message(ui, ERROR_MSG.format(traceback.format_exc()))
