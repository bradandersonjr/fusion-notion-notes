"""
Notion Quick Note - Fusion 360 Add-in
=====================================

A simple Fusion 360 add-in that provides quick access to create new Notion pages
directly from the Fusion 360 interface. This add-in adds a button to the Quick Access
Toolbar (QAT) that opens a new Notion page in your default web browser.

Author: Brad Anderson Jr
Contact: brad@bradandersonjr.com
Version: 0.1.0

Features:
- One-click access to create new Notion pages
- Integrates seamlessly with Fusion 360's Quick Access Toolbar
- Opens Notion in your default web browser for maximum compatibility
- Clean error handling with user-friendly messages

Installation:
1. Copy this add-in folder to your Fusion 360 add-ins directory
2. Start/restart Fusion 360
3. Go to Scripts and Add-Ins dialog
4. Select this add-in and click "Run"
5. The Notion Quick Note button will appear in your Quick Access Toolbar

Usage:
- Click the "Notion Quick Note" button in the Quick Access Toolbar
- A new Notion page will open in your default web browser
- Start taking notes or creating content immediately

Technical Details:
- Uses Fusion 360's Command and Event Handler architecture
- Leverages Python's webbrowser module for cross-platform compatibility
- Follows Fusion 360 add-in best practices for UI integration
- Includes comprehensive error handling and cleanup procedures
"""

import adsk.core, adsk.fusion, traceback, webbrowser

# ============================================================================
# CONSTANTS AND CONFIGURATION
# ============================================================================
ADDIN_NAME = 'Notion Quick Note'  # Display name for the add-in
ADDIN_VERSION = '0.1.0'  # Current version number
ADDIN_AUTHOR = 'Brad Anderson Jr'  # Add-in author
ADDIN_CONTACT = 'brad@bradandersonjr.com'  # Contact information
ERROR_MSG = 'Failed:\n{}'  # Template for error message formatting

# ============================================================================
# GLOBAL VARIABLES
# ============================================================================

# List to hold event handlers to prevent them from being garbage collected
# This is critical in Fusion 360 add-ins - without this, Python's garbage
# collector may clean up event handlers, causing the add-in to stop working
handlers = []

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def show_error_message(ui, error_message):
    """
    Displays an error message dialog to the user in Fusion 360.
    
    This function provides a standardized way to show error messages with
    consistent formatting and branding. It safely handles cases where the
    UI object might be None or invalid.

    Args:
        ui (adsk.core.UserInterface): The Fusion 360 user interface object.
                                    If None, the function will safely exit
                                    without displaying anything.
        error_message (str): The error message text to display to the user.
                           Should be descriptive and user-friendly.
    
    Returns:
        None
        
    Note:
        The message box parameters (0, 0) specify:
        - Message box type: 0 = OKMessageBoxButtonType
        - Icon type: 0 = NoIconMessageBoxIconType
    """
    if ui:
        ui.messageBox(error_message, ADDIN_NAME, 0, 0)

# ============================================================================
# EVENT HANDLER CLASSES
# ============================================================================

class NewNotionPageCommandExecuteHandler(adsk.core.CommandEventHandler):
    """
    Handles the execution of the "New Notion Page" command.
    
    This class is responsible for the actual functionality when the user clicks
    the Notion Quick Note button. It inherits from Fusion 360's CommandEventHandler
    and implements the notify method to define what happens when the command is executed.
    
    The handler simply opens a new Notion page in the user's default web browser
    using Python's webbrowser module for maximum cross-platform compatibility.
    """
    
    def __init__(self, ui):
        """
        Initialize the command execute handler.
        
        Args:
            ui (adsk.core.UserInterface): The Fusion 360 user interface object,
                                        used for displaying error messages if needed.
        """
        super().__init__()
        self.ui = ui  # Store UI reference for error message display
    
    def notify(self, args):
        """
        Called when the user clicks the Notion Quick Note button.
        
        This method is automatically invoked by Fusion 360 when the command
        is executed. It opens a new Notion page in the user's default web browser.
        
        Args:
            args (adsk.core.CommandEventArgs): Event arguments from Fusion 360.
                                             Not used in this implementation but
                                             required by the interface.
        
        Returns:
            None
            
        Note:
            Any exceptions are caught and displayed to the user via a message box
            to prevent the add-in from crashing Fusion 360.
        """
        try:
            # Open new Notion page in the user's default web browser
            # The webbrowser.open_new() function creates a new browser window/tab
            webbrowser.open_new('https://www.notion.so/new')
                
        except Exception as e:
            # If anything goes wrong, show a user-friendly error message
            # The traceback provides detailed error information for debugging
            show_error_message(self.ui, ERROR_MSG.format(traceback.format_exc()))


class NewNotionPageCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """
    Handles the creation and setup of the "New Notion Page" command.
    
    This class is responsible for setting up the command when it's first created
    by Fusion 360. It attaches the execute handler that defines what happens
    when the user actually clicks the button.
    
    This separation of concerns (creation vs execution) is part of Fusion 360's
    event-driven architecture and allows for more complex command behaviors
    like input validation, custom UI dialogs, etc.
    """
    
    def __init__(self, ui):
        """
        Initialize the command created handler.
        
        Args:
            ui (adsk.core.UserInterface): The Fusion 360 user interface object,
                                        used for displaying error messages if needed.
        """
        super().__init__()
        self.ui = ui  # Store UI reference for error message display
    
    def notify(self, args):
        """
        Called when Fusion 360 creates the command definition.
        
        This method is automatically invoked by Fusion 360 when the command
        is being set up. It attaches the execute handler that will be called
        when the user actually clicks the button.
        
        Args:
            args (adsk.core.CommandCreatedEventArgs): Event arguments containing
                                                    the command object to configure.
        
        Returns:
            None
            
        Note:
            The execute handler is added to the global handlers list to prevent
            Python's garbage collector from cleaning it up prematurely.
        """
        try:
            # Get the command object from the event arguments
            command = args.command
            
            # Create the execute handler that will handle button clicks
            onExecute = NewNotionPageCommandExecuteHandler(self.ui)
            
            # Attach the execute handler to the command
            command.execute.add(onExecute)
            
            # Store the handler in the global list to prevent garbage collection
            # This is critical - without this, the handler may be cleaned up
            # and the button will stop working
            handlers.append(onExecute)
            
        except Exception as e:
            # If anything goes wrong during setup, show an error message
            show_error_message(self.ui, ERROR_MSG.format(traceback.format_exc()))



# ============================================================================
# MAIN ADD-IN FUNCTIONS
# ============================================================================

def run(context):
    """
    Entry point for the Fusion 360 add-in when it's started.
    
    This function is automatically called by Fusion 360 when the add-in is loaded.
    It's responsible for setting up the user interface elements and registering
    event handlers. The function creates a button in the Quick Access Toolbar (QAT)
    that users can click to open new Notion pages.
    
    The setup process involves:
    1. Getting references to Fusion 360's application and UI objects
    2. Creating a command definition that describes the button
    3. Setting up event handlers for when the command is created and executed
    4. Adding the button to the Quick Access Toolbar
    
    Args:
        context: Fusion 360 context object (not used in this implementation
                but required by the add-in interface)
    
    Returns:
        None
        
    Note:
        All setup is wrapped in a try-catch block to prevent the add-in from
        crashing Fusion 360 if something goes wrong during initialization.
    """
    try:
        # Get the main Fusion 360 application object and user interface
        app = adsk.core.Application.get()
        ui = app.userInterface

        # Get reference to the Quick Access Toolbar (QAT)
        # This is the toolbar at the top of Fusion 360 with commonly used commands
        qatToolbar = ui.toolbars.itemById('QAT')
        
        # Create a unique command ID by combining the add-in name with 'CmdDef'
        # This ensures our command doesn't conflict with other add-ins
        command_id = f'{ADDIN_NAME}CmdDef'
        
        # Create the command definition that describes our button
        # Parameters: command_id, display_name, tooltip, icon_folder
        newNotionCmdDef = ui.commandDefinitions.addButtonDefinition(
            command_id, 
            ADDIN_NAME, 
            'Create a new Notion page in browser',  # Tooltip text
            './resources'  # Folder containing button icons
        )
        
        # Create the command created handler and attach it to the command definition
        # This handler will be called when Fusion 360 creates the command
        onCommandCreated = NewNotionPageCommandCreatedHandler(ui)
        newNotionCmdDef.commandCreated.add(onCommandCreated)
        
        # Store the handler in our global list to prevent garbage collection
        handlers.append(onCommandCreated)
        
        # Add the command to the Quick Access Toolbar
        # Parameters: command_definition, insert_before_id, is_promoted
        # 'HealthStatusCommand' is used as a reference point for positioning
        # False means the button is not promoted (always visible)
        ctrl = qatToolbar.controls.addCommand(newNotionCmdDef, 'HealthStatusCommand', False)

    except Exception as e:
        # If anything goes wrong during setup, show an error message to the user
        # This prevents the add-in from silently failing
        show_error_message(ui, ERROR_MSG.format(traceback.format_exc()))


def stop(context):
    """
    Cleanup function called when the Fusion 360 add-in is stopped or unloaded.
    
    This function is automatically called by Fusion 360 when the add-in is being
    shut down. It's responsible for cleaning up all UI elements and event handlers
    that were created during the run() function. Proper cleanup is essential to
    prevent memory leaks and ensure that the add-in can be reloaded cleanly.
    
    The cleanup process involves:
    1. Removing the command definition from Fusion 360's command registry
    2. Removing the button from the Quick Access Toolbar
    3. Event handlers are automatically cleaned up when the command is deleted
    
    Args:
        context: Fusion 360 context object (not used in this implementation
                but required by the add-in interface)
    
    Returns:
        None
        
    Note:
        The cleanup code uses defensive programming - it checks if each UI element
        exists before trying to delete it. This prevents errors if the add-in
        is stopped before being fully initialized.
    """
    try:
        # Get the main Fusion 360 application object and user interface
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        # Reconstruct the command ID that was used during setup
        command_id = f'{ADDIN_NAME}CmdDef'
        
        # Find and delete the command definition
        # This removes the command from Fusion 360's command registry
        cmdDef = ui.commandDefinitions.itemById(command_id)
        if cmdDef:
            cmdDef.deleteMe()
        
        # Find and delete the button from the Quick Access Toolbar
        # This removes the visual button that users can see and click
        qatToolbar = ui.toolbars.itemById('QAT')
        cmd = qatToolbar.controls.itemById(command_id)
        if cmd:
            cmd.deleteMe()
               
    except:
        # If anything goes wrong during cleanup, show an error message
        # We use a bare except here because cleanup should be as robust as possible
        # Even if we can't determine the specific error, we want to inform the user
        show_error_message(ui, ERROR_MSG.format(traceback.format_exc()))