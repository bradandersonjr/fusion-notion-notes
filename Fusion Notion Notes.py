"""
Fusion Notion Notes - Autodesk Fusion Add-in
=====================================

A professional Autodesk Fusion add-in that provides quick access to create new Notion pages
directly from the Autodesk Fusion interface. This add-in integrates seamlessly with Fusion's
Quick Access Toolbar (QAT) and provides an elegant HTML-based settings interface.

Author: Brad Anderson Jr
Contact: brad@bradandersonjr.com
Version: 0.6.0

Features:
    - One-click QAT button for instant Notion page creation
    - Configurable settings panel accessible from Scripts menu
    - Dual-mode operation: web browser or desktop app
    - Persistent configuration with JSON storage
    - No authentication required - uses public Notion URLs
    - Automatic fallback when desktop app is unavailable
    - Comprehensive error handling with user-friendly messages

Installation:
    1. Copy this add-in folder to your Autodesk Fusion add-ins directory
    2. Start or restart Autodesk Fusion
    3. Open the Scripts and Add-Ins dialog
    4. Select this add-in and click "Run"
    5. The Fusion Notion Notes button will appear in your Quick Access Toolbar

Usage:
    - Click the "Fusion Notion Notes" button in the QAT to create a new Notion page
    - Configure your preferred database URL and open method in the settings
    - Settings automatically save and persist across sessions
    - Desktop app mode automatically falls back to web browser if app is not installed

Configuration:
    - Database URL: The Notion database where new pages will be created
    - Open Method: Choose between 'web' (browser) or 'desktop' (Notion app)
    - All settings are stored locally in notion_config.json
    - No API keys or authentication tokens required

Technical Details:
    - Built on Autodesk Fusion's Command and Event Handler architecture
    - Uses HTML5-based palette for modern, responsive UI
    - Cross-platform browser integration via Python's webbrowser module
    - Protocol handler detection for notion:// URL scheme
    - JSON-based configuration management
    - Follows Autodesk Fusion add-in best practices
    - Comprehensive cleanup procedures to prevent memory leaks
"""

import adsk.core
import adsk.fusion
import traceback
import webbrowser
import json
import os
import subprocess
import platform
from typing import Optional, Dict, Any, Tuple

# ============================================================================
# CONSTANTS AND CONFIGURATION
# ============================================================================

# Add-in metadata
ADDIN_NAME = 'Fusion Notion Notes'
ADDIN_VERSION = '0.6.0'
ADDIN_AUTHOR = 'Brad Anderson Jr'
ADDIN_CONTACT = 'brad@bradandersonjr.com'

# File and path constants
CONFIG_FILENAME = 'notion_config.json'

# UI element identifiers
PALETTE_ID = 'FusionNotionNotesPalette'
SETTINGS_CMD_ID = 'FusionNotionNotesSettings'

# Default configuration values
DEFAULT_NOTION_URL = 'https://www.notion.so/new'
DEFAULT_OPEN_METHOD = 'web'

# UI dimensions
PALETTE_WIDTH = 720
PALETTE_HEIGHT = 900

# Message box constants
MSG_BOX_OK_ONLY = 0
MSG_BOX_INFO_ICON = 0

# Error message template
ERROR_MSG_TEMPLATE = 'Failed:\n{}'

# Registry paths (Windows only)
WINDOWS_REGISTRY_NOTION_KEY = 'HKEY_CLASSES_ROOT\\notion'
REGISTRY_QUERY_TIMEOUT = 2  # seconds

# ============================================================================
# GLOBAL VARIABLES
# ============================================================================

# Event handler storage to prevent garbage collection
handlers = []

# Palette instance
palette = None

# ============================================================================
# CONFIGURATION MANAGEMENT
# ============================================================================

def get_config_path() -> str:
    """
    Retrieves the absolute path to the configuration file.

    The configuration file is stored in the same directory as this add-in script,
    ensuring it persists across Fusion sessions and is specific to this add-in.

    Returns:
        str: Absolute path to the notion_config.json file

    Example:
        >>> config_path = get_config_path()
        >>> print(config_path)
        'C:/Users/User/AppData/Roaming/Autodesk/.../notion_config.json'
    """
    addin_dir = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(addin_dir, CONFIG_FILENAME)


def create_default_config() -> Dict[str, str]:
    """
    Creates and returns a default configuration dictionary.

    This configuration is used when no existing config file is found or when
    loading the config file fails. It provides sensible defaults for first-time
    users.

    Returns:
        Dict[str, str]: Default configuration with database URL and open method

    Example:
        >>> config = create_default_config()
        >>> print(config['default_open_method'])
        'web'
    """
    return {
        'database_url': DEFAULT_NOTION_URL,
        'default_open_method': DEFAULT_OPEN_METHOD
    }


def load_config() -> Dict[str, str]:
    """
    Loads the configuration from the JSON file.

    This function attempts to read the configuration file. If the file doesn't
    exist or cannot be parsed, it creates a new default configuration file
    and returns the default values.

    Returns:
        Dict[str, str]: Configuration dictionary containing 'database_url' and
                       'default_open_method' keys

    Raises:
        No exceptions are raised; errors are handled gracefully with defaults

    Example:
        >>> config = load_config()
        >>> url = config.get('database_url')
        >>> method = config.get('default_open_method')
    """
    config_path = get_config_path()

    # Create default config if file doesn't exist
    if not os.path.exists(config_path):
        default_config = create_default_config()
        save_config(default_config)
        return default_config

    # Attempt to load existing config
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        # Validate that required keys exist
        if 'database_url' not in config or 'default_open_method' not in config:
            raise ValueError('Invalid config format')

        return config

    except (json.JSONDecodeError, IOError, ValueError) as e:
        # If loading fails, return default config
        # Log the error silently and proceed with defaults
        return create_default_config()


def save_config(config: Dict[str, str]) -> bool:
    """
    Saves the configuration to the JSON file.

    Writes the configuration dictionary to the JSON file with proper formatting
    for human readability. The file is created if it doesn't exist.

    Args:
        config (Dict[str, str]): Configuration dictionary to save

    Returns:
        bool: True if save was successful, False otherwise

    Example:
        >>> config = {'database_url': 'https://...', 'default_open_method': 'web'}
        >>> success = save_config(config)
        >>> print(success)
        True
    """
    config_path = get_config_path()

    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True

    except (IOError, TypeError) as e:
        # Save failed - return False to indicate failure
        return False


# ============================================================================
# NOTION URL GENERATION
# ============================================================================

def get_notion_url(protocol: str = 'https') -> str:
    """
    Generates the appropriate Notion URL based on configuration and protocol.

    This function reads the user's configured database URL and converts it to
    the appropriate format based on the requested protocol. For the 'notion'
    protocol (desktop app), it converts https:// URLs to notion:// URLs.

    Args:
        protocol (str): The URL protocol to use. Options are:
                       - 'https': For web browser access
                       - 'notion': For desktop app access

    Returns:
        str: The formatted Notion URL ready to be opened

    Example:
        >>> web_url = get_notion_url('https')
        >>> print(web_url)
        'https://www.notion.so/database/...'

        >>> desktop_url = get_notion_url('notion')
        >>> print(desktop_url)
        'notion://www.notion.so/database/...'
    """
    config = load_config()
    database_url = config.get('database_url', '').strip()

    # Return configured database URL or default based on protocol
    if database_url:
        # User has configured a custom database URL
        if protocol == 'notion':
            # Convert web URL to desktop app protocol
            if database_url.startswith('https://'):
                return database_url.replace('https://', 'notion://', 1)
            elif database_url.startswith('http://'):
                return database_url.replace('http://', 'notion://', 1)
            # If already notion://, return as-is
            return database_url
        else:
            # Return web URL as-is
            return database_url
    else:
        # No database configured - use default new page URL
        if protocol == 'notion':
            return 'notion://www.notion.so/new'
        else:
            return DEFAULT_NOTION_URL


# ============================================================================
# NOTION DESKTOP APP DETECTION
# ============================================================================

def check_notion_protocol_handler() -> bool:
    """
    Checks if the Notion desktop app protocol handler (notion://) is available.

    This function performs platform-specific checks to determine if the Notion
    desktop application is installed and has registered its protocol handler:

    - Windows: Queries the registry for the notion:// protocol registration
    - macOS/Linux: Assumes protocol might exist (webbrowser will handle errors)

    Returns:
        bool: True if protocol handler is detected, False otherwise

    Platform-specific behavior:
        - Windows: Uses 'reg query' to check HKEY_CLASSES_ROOT\notion
        - macOS: Returns True (protocol detection not easily available)
        - Linux: Returns True (protocol detection not easily available)

    Example:
        >>> has_desktop_app = check_notion_protocol_handler()
        >>> if has_desktop_app:
        ...     print("Notion desktop app is installed")
        ... else:
        ...     print("Notion desktop app not found")
    """
    try:
        if platform.system() == 'Windows':
            # Windows: Check registry for notion:// protocol handler
            try:
                # Query the Windows registry for the Notion protocol key
                result = subprocess.run(
                    ['reg', 'query', WINDOWS_REGISTRY_NOTION_KEY, '/ve'],
                    capture_output=True,
                    timeout=REGISTRY_QUERY_TIMEOUT,
                    creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                )
                # Return True if registry key exists (returncode == 0)
                return result.returncode == 0

            except (subprocess.TimeoutExpired, subprocess.SubprocessError, OSError):
                # If registry check fails, assume protocol doesn't exist
                return False
        else:
            # macOS/Linux: Cannot easily check for protocol handler
            # Return True and let webbrowser module handle any errors
            return True

    except Exception:
        # If anything unexpected happens, return False to be safe
        return False


# ============================================================================
# NOTION OPENING WITH FALLBACK
# ============================================================================

def open_notion_with_fallback(protocol: str = 'https', ui: Optional[adsk.core.UserInterface] = None) -> bool:
    """
    Opens Notion URL with automatic fallback to web browser if desktop app fails.

    This function provides a robust way to open Notion pages with intelligent
    fallback behavior. When the desktop app is requested but unavailable, it
    automatically falls back to the web browser and notifies the user.

    Args:
        protocol (str): The protocol to use for opening Notion:
                       - 'notion': Attempt to use desktop app (with fallback)
                       - 'https': Use web browser directly
        ui (Optional[adsk.core.UserInterface]): User interface object for
                                                displaying messages. Can be None.

    Returns:
        bool: True if opened successfully with requested method,
              False if fallback was used or operation failed

    Side Effects:
        - Opens a new browser tab or desktop app window
        - May display a message box to user if fallback occurs

    Example:
        >>> # Try to open in desktop app
        >>> success = open_notion_with_fallback('notion', ui)
        >>> if not success:
        ...     print("Fell back to web browser")

        >>> # Open directly in web browser
        >>> open_notion_with_fallback('https', ui)
    """
    try:
        url = get_notion_url(protocol=protocol)

        if protocol == 'notion':
            # Desktop app mode - check if protocol handler exists
            if not check_notion_protocol_handler():
                # Protocol handler not found - fall back to web browser
                return _fallback_to_web_browser(
                    'Notion desktop app not found. Opened in web browser instead.',
                    ui
                )

            # Protocol handler exists - try to open desktop app
            try:
                webbrowser.open(url)
                return True

            except Exception:
                # Opening desktop app failed - fall back to web browser
                return _fallback_to_web_browser(
                    'Could not open Notion desktop app. Opened in web browser instead.',
                    ui
                )
        else:
            # Web browser mode - open directly
            webbrowser.open_new(url)
            return True

    except Exception as e:
        # Last resort fallback - try web browser one final time
        return _fallback_to_web_browser(
            'Could not open Notion desktop app. Opened in web browser instead.',
            ui,
            show_message=(protocol == 'notion')
        )


def _fallback_to_web_browser(message: str, ui: Optional[adsk.core.UserInterface], show_message: bool = True) -> bool:
    """
    Internal helper function to fall back to web browser with user notification.

    This function is called when the desktop app cannot be opened. It opens the
    web browser as a fallback and optionally displays a message to the user.

    Args:
        message (str): The message to display to the user
        ui (Optional[adsk.core.UserInterface]): User interface for message display
        show_message (bool): Whether to show the message box to the user

    Returns:
        bool: False to indicate fallback was used (not the requested method)

    Side Effects:
        - Opens web browser with Notion URL
        - May display a message box to the user
    """
    try:
        web_url = get_notion_url(protocol='https')
        webbrowser.open_new(web_url)

        # Show informational message to user if UI is available
        if ui and show_message:
            ui.messageBox(
                message,
                ADDIN_NAME,
                MSG_BOX_OK_ONLY,
                MSG_BOX_INFO_ICON
            )

        return False  # Return False because we used fallback, not requested method

    except Exception:
        # If even web browser fails, return False silently
        return False


# ============================================================================
# USER INTERFACE HELPERS
# ============================================================================

def show_error_message(ui: Optional[adsk.core.UserInterface], error_message: str) -> None:
    """
    Displays an error message dialog to the user.

    This function provides a consistent way to show error messages throughout
    the add-in. It safely handles cases where the UI object is not available.

    Args:
        ui (Optional[adsk.core.UserInterface]): User interface object for
                                                displaying the message. Can be None.
        error_message (str): The error message to display to the user

    Side Effects:
        - Displays a modal message box with the error message

    Example:
        >>> try:
        ...     risky_operation()
        ... except Exception as e:
        ...     show_error_message(ui, f"Operation failed: {str(e)}")
    """
    if ui:
        ui.messageBox(error_message, ADDIN_NAME, MSG_BOX_OK_ONLY, MSG_BOX_INFO_ICON)


def send_config_to_palette(palette_instance: adsk.core.Palette) -> None:
    """
    Sends the current configuration to the HTML palette.

    This function updates the palette's UI with the current configuration values
    by sending a JSON message to the HTML interface. It's used to populate the
    settings form when the palette is opened.

    Args:
        palette_instance (adsk.core.Palette): The palette to send config to

    Side Effects:
        - Sends JSON data to the HTML palette via sendInfoToHTML

    Example:
        >>> send_config_to_palette(palette)
        # Palette UI updates with current configuration
    """
    if palette_instance and palette_instance.isVisible:
        try:
            config = load_config()
            config_data = json.dumps({
                'action': 'setConfig',
                'databaseUrl': config.get('database_url', ''),
                'defaultMethod': config.get('default_open_method', DEFAULT_OPEN_METHOD)
            })
            palette_instance.sendInfoToHTML('setConfig', config_data)

        except Exception:
            # Silently fail if palette is not ready or communication fails
            pass


# ============================================================================
# EVENT HANDLERS - PALETTE
# ============================================================================

class PaletteCommandHandler(adsk.core.HTMLEventHandler):
    """
    Handles events and messages from the HTML palette.

    This handler receives messages sent from the JavaScript code in the palette's
    HTML file. It processes different actions like getting/saving configuration,
    and opening URLs.

    Attributes:
        ui (adsk.core.UserInterface): Reference to the Fusion UI for displaying messages

    Supported Actions:
        - 'getConfig': Requests current configuration to populate settings form
        - 'savePreferences': Saves user's updated preferences to config file
        - 'openNotionForUrl': Opens Notion website to help user get database URL
        - 'openUrl': Opens an arbitrary URL in the browser
    """

    def __init__(self, ui: adsk.core.UserInterface):
        """
        Initialize the palette command handler.

        Args:
            ui (adsk.core.UserInterface): The Fusion user interface object
        """
        super().__init__()
        self.ui = ui

    def notify(self, args: adsk.core.HTMLEventArgs) -> None:
        """
        Called when the HTML palette sends a message to the add-in.

        This method processes incoming messages from the palette's JavaScript
        code and performs the appropriate actions based on the message type.

        Args:
            args (adsk.core.HTMLEventArgs): Event arguments containing the action
                                           and data from the HTML palette

        Side Effects:
            - May modify configuration file
            - May open browser windows
            - May send data back to palette
        """
        try:
            htmlArgs = adsk.core.HTMLEventArgs.cast(args)
            action = htmlArgs.action

            if action == 'getConfig':
                # Return current configuration to the palette
                config = load_config()
                return_data = json.dumps({
                    'action': 'setConfig',
                    'databaseUrl': config.get('database_url', ''),
                    'defaultMethod': config.get('default_open_method', DEFAULT_OPEN_METHOD)
                })
                htmlArgs.returnData = return_data

                # Also send via sendInfoToHTML as backup method
                try:
                    global palette
                    if palette and palette.isVisible:
                        palette.sendInfoToHTML('setConfig', return_data)
                except Exception:
                    pass  # Silently fail if palette not ready

            elif action == 'savePreferences':
                # Save user's updated preferences from the settings panel
                data = json.loads(htmlArgs.data) if htmlArgs.data else {}
                config = load_config()

                # Extract and validate new settings
                database_url = data.get('databaseUrl', '').strip()
                default_method = data.get('defaultMethod', DEFAULT_OPEN_METHOD)

                # Update configuration
                config['database_url'] = database_url
                config['default_open_method'] = default_method

                # Persist to file
                save_config(config)

            elif action == 'openNotionForUrl':
                # Open Notion website to help user navigate and get database URL
                webbrowser.open_new('https://www.notion.so')

            elif action == 'openUrl':
                # Open any URL in the user's browser (for help links, etc.)
                url = htmlArgs.data if htmlArgs.data else ''
                if url:
                    webbrowser.open_new(url)

        except Exception as e:
            # Display error to user if something goes wrong
            error_msg = ERROR_MSG_TEMPLATE.format(traceback.format_exc())
            show_error_message(self.ui, error_msg)


class PaletteClosedHandler(adsk.core.UserInterfaceGeneralEventHandler):
    """
    Handles events when the palette is closed by the user.

    This handler is notified when the palette's close button is clicked or when
    it's programmatically closed. Currently a placeholder for future functionality.

    Attributes:
        ui (adsk.core.UserInterface): Reference to the Fusion UI
    """

    def __init__(self, ui: adsk.core.UserInterface):
        """
        Initialize the palette closed handler.

        Args:
            ui (adsk.core.UserInterface): The Fusion user interface object
        """
        super().__init__()
        self.ui = ui

    def notify(self, args: adsk.core.UserInterfaceGeneralEventArgs) -> None:
        """
        Called when the palette is closed.

        Currently a no-op as the configuration is sent when palette reopens.

        Args:
            args (adsk.core.UserInterfaceGeneralEventArgs): Event arguments
        """
        try:
            # When palette is closed, no action needed
            # Config will be sent fresh when it's reopened
            pass
        except Exception:
            # Silently fail - closing should always succeed
            pass


# ============================================================================
# EVENT HANDLERS - COMMANDS
# ============================================================================

class NotionQuickOpenHandler(adsk.core.CommandEventHandler):
    """
    Handles the execution of the Quick Access Toolbar button click.

    This handler is triggered when the user clicks the "Fusion Notion Notes"
    button in the QAT. It reads the user's preferred open method and opens
    Notion accordingly (desktop app or web browser).

    Attributes:
        ui (adsk.core.UserInterface): Reference to the Fusion UI
    """

    def __init__(self, ui: adsk.core.UserInterface):
        """
        Initialize the quick open handler.

        Args:
            ui (adsk.core.UserInterface): The Fusion user interface object
        """
        super().__init__()
        self.ui = ui

    def notify(self, args: adsk.core.CommandEventArgs) -> None:
        """
        Called when the user clicks the Fusion Notion Notes button in the QAT.

        This method loads the user's configuration and opens Notion using their
        preferred method (desktop app or web browser) with automatic fallback.

        Args:
            args (adsk.core.CommandEventArgs): Command execution event arguments

        Side Effects:
            - Opens Notion in web browser or desktop app
            - May display a message if fallback occurs
        """
        try:
            config = load_config()
            database_url = config.get('database_url', '').strip()
            default_method = config.get('default_open_method', DEFAULT_OPEN_METHOD)

            if not database_url:
                # No database URL configured - show settings palette instead
                global palette
                if palette:
                    palette.isVisible = True
                    send_config_to_palette(palette)
                else:
                    settings_handler = NotionSettingsHandler(self.ui)
                    palette = self.ui.palettes.itemById(PALETTE_ID)
                    if not palette:
                        palette = settings_handler._create_palette()
                    else:
                        palette.isVisible = True
                        send_config_to_palette(palette)
                return

            if default_method == 'desktop':
                # Try desktop app with automatic fallback to web browser
                open_notion_with_fallback(protocol='notion', ui=self.ui)
            else:
                # Open directly in web browser
                open_notion_with_fallback(protocol='https', ui=self.ui)

        except Exception as e:
            # Display error message if something goes wrong
            error_msg = ERROR_MSG_TEMPLATE.format(traceback.format_exc())
            show_error_message(self.ui, error_msg)


class NotionSettingsHandler(adsk.core.CommandEventHandler):
    """
    Handles showing and hiding the settings palette.

    This handler is triggered when the user clicks the "Fusion Notion Notes Settings"
    command in the Scripts menu. It manages the palette lifecycle including creation,
    visibility toggling, and configuration updates.

    Attributes:
        ui (adsk.core.UserInterface): Reference to the Fusion UI
    """

    def __init__(self, ui: adsk.core.UserInterface):
        """
        Initialize the settings handler.

        Args:
            ui (adsk.core.UserInterface): The Fusion user interface object
        """
        super().__init__()
        self.ui = ui

    def notify(self, args: adsk.core.CommandEventArgs) -> None:
        """
        Called when the user clicks the settings command in the Scripts menu.

        This method manages the settings palette by toggling its visibility if it
        exists, or creating it if it doesn't. When shown, it sends the current
        configuration to populate the form.

        Args:
            args (adsk.core.CommandEventArgs): Command execution event arguments

        Side Effects:
            - Creates palette HTML window if needed
            - Toggles palette visibility
            - Sends configuration data to palette
            - Registers event handlers for palette communication
        """
        try:
            global palette

            if palette:
                # Toggle palette visibility if it already exists
                was_visible = palette.isVisible
                palette.isVisible = not was_visible

                # Send fresh config when showing palette
                if palette.isVisible:
                    send_config_to_palette(palette)
            else:
                # Create the palette for the first time
                palette = self.ui.palettes.itemById(PALETTE_ID)

                if not palette:
                    # Palette doesn't exist - create it
                    palette = self._create_palette()
                else:
                    # Palette exists but was hidden - show it
                    palette.isVisible = True
                    send_config_to_palette(palette)

        except Exception as e:
            # Display error message if palette creation/toggle fails
            error_msg = ERROR_MSG_TEMPLATE.format(traceback.format_exc())
            show_error_message(self.ui, error_msg)

    def _create_palette(self) -> adsk.core.Palette:
        """
        Creates and configures the HTML-based settings palette.

        This internal method handles the complete palette creation process including:
        - Loading the HTML file
        - Setting up event handlers
        - Configuring palette properties (size, docking, etc.)
        - Sending initial configuration data

        Returns:
            adsk.core.Palette: The newly created palette instance

        Side Effects:
            - Creates new palette window
            - Registers event handlers (stored in global handlers list)
            - Sends initial configuration to HTML
        """
        # Get the HTML file path
        addin_dir = os.path.dirname(os.path.realpath(__file__))
        
        # Find the HTML file (case-insensitive on Windows)
        html_file = None
        possible_names = ['Palette.html', 'palette.html', 'PALETTE.HTML']
        for name in possible_names:
            test_path = os.path.join(addin_dir, name)
            if os.path.exists(test_path):
                html_file = test_path
                break
        
        # Fallback to default if not found
        if not html_file:
            html_file = os.path.join(addin_dir, 'Palette.html')

        # Create a temporary HTML file with injected config
        temp_html_file = os.path.join(addin_dir, 'Palette_temp.html')

        try:
            # Read the template HTML
            with open(html_file, 'r', encoding='utf-8') as f:
                html_content = f.read()

            # Load current config
            config = load_config()
            config_json = json.dumps({
                'databaseUrl': config.get('database_url', ''),
                'defaultMethod': config.get('default_open_method', DEFAULT_OPEN_METHOD)
            })

            # Inject config as a script tag right after <head>
            injection = f'''<head>
    <script>
        window.FUSION_NOTION_CONFIG = {config_json};
    </script>'''

            html_content = html_content.replace('<head>', injection, 1)

            # Write temporary HTML file
            with open(temp_html_file, 'w', encoding='utf-8') as f:
                f.write(html_content)

            html_file_url = temp_html_file.replace('\\', '/')

        except Exception as e:
            # If injection fails, fall back to original HTML
            # But still try to send config via sendInfoToHTML
            html_file_url = html_file.replace('\\', '/')

        # Create the palette with configured properties
        new_palette = self.ui.palettes.add(
            PALETTE_ID,
            'Fusion Notion Notes Settings',
            html_file_url,
            True,  # Show palette immediately
            True,  # Show close button
            True,  # Can be resized by user
            PALETTE_WIDTH,
            PALETTE_HEIGHT
        )

        # Register HTML event handler for palette communication
        on_html = PaletteCommandHandler(self.ui)
        new_palette.incomingFromHTML.add(on_html)
        handlers.append(on_html)

        # Register closed event handler
        on_closed = PaletteClosedHandler(self.ui)
        new_palette.closed.add(on_closed)
        handlers.append(on_closed)

        # Set palette to float (not docked to any side)
        new_palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateFloating

        # Send initial configuration to populate the form
        send_config_to_palette(new_palette)

        return new_palette


class NotionQuickOpenCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """
    Handles the creation of the Quick Access Toolbar button command.

    This handler is called when the QAT button command definition is created.
    It's responsible for attaching the execute handler that runs when the button
    is clicked.

    Attributes:
        ui (adsk.core.UserInterface): Reference to the Fusion UI
    """

    def __init__(self, ui: adsk.core.UserInterface):
        """
        Initialize the command created handler for QAT button.

        Args:
            ui (adsk.core.UserInterface): The Fusion user interface object
        """
        super().__init__()
        self.ui = ui

    def notify(self, args: adsk.core.CommandCreatedEventArgs) -> None:
        """
        Called when the QAT button command is created.

        Attaches the execute handler that will run when the button is clicked.

        Args:
            args (adsk.core.CommandCreatedEventArgs): Command creation event arguments

        Side Effects:
            - Attaches execute handler to command
            - Stores handler in global list to prevent garbage collection
        """
        try:
            command = args.command

            # Attach execute handler for button clicks
            on_execute = NotionQuickOpenHandler(self.ui)
            command.execute.add(on_execute)
            handlers.append(on_execute)

        except Exception as e:
            # Display error message if handler attachment fails
            error_msg = ERROR_MSG_TEMPLATE.format(traceback.format_exc())
            show_error_message(self.ui, error_msg)


class NotionSettingsCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """
    Handles the creation of the Settings command.

    This handler is called when the settings command definition is created.
    It's responsible for attaching the execute handler that runs when the
    settings menu item is clicked.

    Attributes:
        ui (adsk.core.UserInterface): Reference to the Fusion UI
    """

    def __init__(self, ui: adsk.core.UserInterface):
        """
        Initialize the command created handler for settings command.

        Args:
            ui (adsk.core.UserInterface): The Fusion user interface object
        """
        super().__init__()
        self.ui = ui

    def notify(self, args: adsk.core.CommandCreatedEventArgs) -> None:
        """
        Called when the settings command is created.

        Attaches the execute handler that will run when the command is executed.

        Args:
            args (adsk.core.CommandCreatedEventArgs): Command creation event arguments

        Side Effects:
            - Attaches execute handler to command
            - Stores handler in global list to prevent garbage collection
        """
        try:
            command = args.command

            # Attach execute handler for command execution
            on_execute = NotionSettingsHandler(self.ui)
            command.execute.add(on_execute)
            handlers.append(on_execute)

        except Exception as e:
            # Display error message if handler attachment fails
            error_msg = ERROR_MSG_TEMPLATE.format(traceback.format_exc())
            show_error_message(self.ui, error_msg)


# ============================================================================
# ADD-IN LIFECYCLE FUNCTIONS
# ============================================================================

def run(context: Dict[str, Any]) -> None:
    """
    Entry point for the Autodesk Fusion add-in.

    This function is called by Fusion when the add-in is loaded. It sets up
    all UI elements including:
    - Quick Access Toolbar button for quick Notion access
    - Settings command in the Scripts menu
    - Event handlers for all commands

    The function is designed to be idempotent and handles errors gracefully.

    Args:
        context (Dict[str, Any]): Context dictionary provided by Fusion

    Side Effects:
        - Creates command definitions
        - Adds buttons to QAT and Scripts menu
        - Registers event handlers (stored in global handlers list)

    Example:
        This function is called automatically by Fusion, not by user code.
    """
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        # ====================================================================
        # QUICK ACCESS TOOLBAR BUTTON
        # ====================================================================

        # Get reference to the Quick Access Toolbar
        qat_toolbar = ui.toolbars.itemById('QAT')

        # Create command definition for QAT button
        qat_command_id = f'{ADDIN_NAME}CmdDef'
        notion_quick_open_cmd = ui.commandDefinitions.addButtonDefinition(
            qat_command_id,
            ADDIN_NAME,
            'Create a new Notion page',
            './resources'  # Icon directory
        )

        # Attach command created handler
        on_quick_open_created = NotionQuickOpenCommandCreatedHandler(ui)
        notion_quick_open_cmd.commandCreated.add(on_quick_open_created)
        handlers.append(on_quick_open_created)

        # Add button to Quick Access Toolbar (after Health Status if possible)
        qat_toolbar.controls.addCommand(notion_quick_open_cmd, 'HealthStatusCommand', False)

        # ====================================================================
        # SETTINGS COMMAND (SCRIPTS MENU)
        # ====================================================================

        # Create command definition for settings
        notion_settings_cmd = ui.commandDefinitions.addButtonDefinition(
            SETTINGS_CMD_ID,
            'Fusion Notion Notes Settings',
            'Configure Notion database and default open method',
            './resources'  # Icon directory
        )

        # Attach command created handler
        on_settings_created = NotionSettingsCommandCreatedHandler(ui)
        notion_settings_cmd.commandCreated.add(on_settings_created)
        handlers.append(on_settings_created)

        # Add settings command to ADD-INS panel in UTILITIES toolbar
        workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        if workspace:
            addins_panel = workspace.toolbarPanels.itemById('SolidScriptsAddinsPanel')
            if addins_panel:
                control = addins_panel.controls.addCommand(notion_settings_cmd)

                # Ensure control is not a dropdown
                if control and hasattr(control, 'isDropDown'):
                    control.isDropDown = False

    except Exception as e:
        # Display error message if add-in initialization fails
        if ui:
            error_msg = ERROR_MSG_TEMPLATE.format(traceback.format_exc())
            show_error_message(ui, error_msg)


def stop(context: Dict[str, Any]) -> None:
    """
    Cleanup function called when the add-in is stopped or unloaded.

    This function performs comprehensive cleanup of all UI elements and event
    handlers created by the add-in. It ensures that Fusion returns to its
    original state with no remnants of the add-in left behind.

    Cleanup includes:
    - Removing and deleting the settings palette
    - Removing QAT button and its command definition
    - Removing settings command from Scripts menu
    - Clearing all event handlers

    Args:
        context (Dict[str, Any]): Context dictionary provided by Fusion

    Side Effects:
        - Deletes UI controls and command definitions
        - Clears global event handler list
        - Closes and deletes palette if open

    Example:
        This function is called automatically by Fusion, not by user code.
    """
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        # ====================================================================
        # CLEANUP PALETTE
        # ====================================================================

        global palette
        if palette:
            palette.deleteMe()
            palette = None

        # Clean up temporary HTML file
        addin_dir = os.path.dirname(os.path.realpath(__file__))
        temp_html_file = os.path.join(addin_dir, 'Palette_temp.html')
        try:
            if os.path.exists(temp_html_file):
                os.remove(temp_html_file)
        except Exception:
            pass  # Silently fail if cleanup fails

        # ====================================================================
        # CLEANUP QUICK ACCESS TOOLBAR BUTTON
        # ====================================================================

        qat_toolbar = ui.toolbars.itemById('QAT')
        qat_command_id = f'{ADDIN_NAME}CmdDef'

        # Remove command definition
        cmd_def = ui.commandDefinitions.itemById(qat_command_id)
        if cmd_def:
            cmd_def.deleteMe()

        # Remove control from QAT
        cmd_control = qat_toolbar.controls.itemById(qat_command_id)
        if cmd_control:
            cmd_control.deleteMe()

        # ====================================================================
        # CLEANUP SETTINGS COMMAND
        # ====================================================================

        # Remove settings command definition
        settings_cmd_def = ui.commandDefinitions.itemById(SETTINGS_CMD_ID)
        if settings_cmd_def:
            settings_cmd_def.deleteMe()

        # Remove from ADD-INS panel
        workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        if workspace:
            addins_panel = workspace.toolbarPanels.itemById('SolidScriptsAddinsPanel')
            if addins_panel:
                settings_control = addins_panel.controls.itemById(SETTINGS_CMD_ID)
                if settings_control:
                    settings_control.deleteMe()

        # ====================================================================
        # CLEANUP EVENT HANDLERS
        # ====================================================================

        # Clear all handlers to allow garbage collection
        handlers.clear()

    except Exception:
        # Silently fail during cleanup to avoid blocking Fusion shutdown
        # Log to Fusion's text commands window if available
        if ui:
            error_msg = ERROR_MSG_TEMPLATE.format(traceback.format_exc())
            show_error_message(ui, error_msg)
