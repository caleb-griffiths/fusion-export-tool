import adsk.core
import os
from ...lib import fusionAddInUtils as futil
from ... import config
app = adsk.core.Application.get()
ui = app.userInterface


# TODO *** Specify the command identity information. ***
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_cmdDialog'
CMD_NAME = 'EXPORT ALL TO STEP'
CMD_Description = 'Export all configurations in this design to STEP files'

# Specify that the command will be promoted to the panel.
IS_PROMOTED = True

# TODO *** Define the location where the command button will be created. ***
# This is done by specifying the workspace, the tab, and the panel, and the 
# command it will be inserted beside. Not providing the command to position it
# will insert it at the end.
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidScriptsAddinsPanel'
COMMAND_BESIDE_ID = 'ScriptsManagerCommand'

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []


# Executed when add-in is run.
def start():
    # Create a command Definition.
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    # Define an event handler for the command created event. It will be called when the button is clicked.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******** Add a button into the UI so the user can run the command. ********
    # Get the target workspace the button will be created in.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get the panel the button will be created in.
    panel = workspace.toolbarPanels.itemById(PANEL_ID)

    # Create the button command control in the UI after the specified existing command.
    control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)

    # Specify if the command is promoted to the main toolbar. 
    control.isPromoted = IS_PROMOTED


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()


# Function that is called when a user clicks the corresponding button in the UI.
# This defines the contents of the command dialog and connects to the command related events.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Created Event')

    # https://help.autodesk.com/view/fusion360/ENU/?contextId=CommandInputs
    inputs = args.command.commandInputs

    # TODO Define the dialog for your command by adding different inputs to the command.

    inputs.addStringValueInput('author_input', 'Author', 'Caleb Griffiths')
    inputs.addStringValueInput('organization_input', 'Organization', 'The Metal Company')
    inputs.addStringValueInput('authorization_input', 'Authorization', 'S. Fisher & Sons Ltd T/A The Metal Company')
    inputs.addBoolValueInput(
        'select_folder_input',
        'Export Folder',
        False,
        '',
        True
    )

    # TODO Connect to the events that are needed by this command.
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(args.command.validateInputs, command_validate_input, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)


# This event handler is called when the user clicks the OK button in the command dialog or 
# is immediately called after the created event not command inputs were created for the dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    futil.log(f'{CMD_NAME} Command Execute Event')

    inputs = args.command.commandInputs

    author_input = inputs.itemById('author_input')
    organization_input = inputs.itemById('organization_input')
    authorization_input = inputs.itemById('authorization_input')
    folder_input = inputs.itemById('select_folder_input')

    author = author_input.value
    organization = organization_input.value
    authorization = authorization_input.value
    export_folder = folder_input.text

    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)

    if not design:
        ui.messageBox('No active Fusion design found...')
        return
    
    config_table = design.configurationTopTable

    if not config_table:
        ui.messageBox('No configurations found...')
        return
    
    if not export_folder:
        ui.messageBox('Please select an export location first.')
        return
    
    export_manager = design.exportManager
    exported_files = []

    progress_dialog = ui.createProgressDialog()
    progress_dialog.cancelButtonShown = True
    progress_dialog.isBackgroundTranslucent = False
    progress_dialog.show(
        'Exporting Configurations',
        'Exporting %v of %m Configurations...',
        0,
        config_table.rows.count
    )

    for i in range(config_table.rows.count):
        row = config_table.rows.item(i)

        current_number = i + 1
        total_count = config_table.rows.count
        percent_done = int((current_number / total_count) * 100)

        progress_dialog.message = (
            'Exporting ' + str(current_number) + ' of ' + str(total_count) + 
            ' Configurations (' + str(percent_done) + '%) ...\n\n' +
            'Exporting: ' + row.name + '.step'
        )

        progress_dialog.progressValue = current_number

        row.activate()

        file_name = row.name + '.step'
        file_path = export_folder + '\\' + file_name

        step_options = export_manager.createSTEPExportOptions(file_path)
        export_manager.execute(step_options)

        with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
            step_text = file.read()

        step_text = step_text.replace(
            "/* author */ (''),",
            f"/* author */ ('{author}'),"
        )

        step_text = step_text.replace(
            "/* organization */ (''),",
            f"/* organization */ ('{organization}'),"
        )

        step_text = step_text.replace(
            "/* authorisation */ '');",
            f"/* authorisation */ '{authorization}');"
        )

        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(step_text)

        exported_files.append(file_name)

    ui.messageBox(
        'Configurations Exported: ' + str(len(exported_files)) + '\n\n' +
        '\n'.join(exported_files)
    )

# This event handler is called when the command needs to compute a new preview in the graphics window.
def command_preview(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Preview Event')
    inputs = args.command.commandInputs


# This event handler is called when the user changes anything in the command dialog
# allowing you to modify values of other inputs based on that change.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input

    if changed_input.id == 'select_folder_input':
        folder_dialog = ui.createFolderDialog()
        folder_dialog.title = 'Select STEP Export Folder'

        dialog_result = folder_dialog.showDialog()

        if dialog_result == adsk.core.DialogResults.DialogOK:
            export_folder = folder_dialog.folder

            folder_input = args.inputs.itemById('select_folder_input')
            folder_input.text = export_folder

# This event handler is called when the user interacts with any of the inputs in the dialog
# which allows you to verify that all of the inputs are valid and enables the OK button.
def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Validate Input Event')

    inputs = args.inputs
    
    # Verify the validity of the input values. This controls if the OK button is enabled or not.
    args.areInputsValid = True        

# This event handler is called when the command terminates.
def command_destroy(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Destroy Event')

    global local_handlers
    local_handlers = []
