Option Explicit On
Option Strict Off

Imports System
Imports Microsoft.VisualStudio.CommandBars
Imports Extensibility
Imports EnvDTE
Imports EnvDTE80

Public Class Connect

    Implements IDTExtensibility2
    Implements IDTCommandTarget

    Private _applicationObject As DTE2
    Private _addInInstance As AddIn
    Private _app As DTE

    '''<summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
    Public Sub New()

    End Sub

    '''<summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
    '''<param name='application'>Root object of the host application.</param>
    '''<param name='connectMode'>Describes how the Add-in is being loaded.</param>
    '''<param name='addInInst'>Object representing this Add-in.</param>
    '''<remarks></remarks>
    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
        _applicationObject = CType(application, EnvDTE80.DTE2)
        _addInInstance = CType(addInInst, EnvDTE.AddIn)
        _app = CType(application, EnvDTE.DTE)
        If connectMode = ext_ConnectMode.ext_cm_UISetup Then
            Dim objAddIn As AddIn = CType(addInInst, AddIn)
            Dim CommandObj As Command
            ' Dim objCommandBar As CommandBar

            'If your command no longer appears on the appropriate command bar, or if you would like to re-create the command,
            ' close all instances of Visual Studio .NET, open a command prompt (MS-DOS window), and run the command 'devenv /setup'.
            'objCommandBar = CType(_applicationObject.Commands.AddCommandBar("BRIEF", vsCommandBarType.vsCommandBarTypeMenu, _applicationObject.CommandBars.Item("Tools")), Microsoft.VisualStudio.CommandBars.CommandBar)

            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFHomeKey", "BRIEFHomeKey", "Emulates BRIEF HOME key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFEndKey", "BRIEFEndKey", "Emulates BRIEF END key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFCopy", "BRIEFCopy", "Emulates BRIEF NUM-PLUS key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFCut", "BRIEFCut", "Emulates BRIEF NUM-MINUS key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFPaste", "BRIEFPaste", "Emulates BRIEF INS key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFLineDelete", "BRIEFLineDelete", "Emulates BRIEF ALT-D key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled

            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFDelete", "BRIEFDelete", "Emulates BRIEF DEL key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFPageUp", "BRIEFPageUp", "Emulates BRIEF PG-UP key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFPageDown", "BRIEFPageDown", "Emulates BRIEF PG-DOWN key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFUndo", "BRIEFUndo", "Emulates BRIEF * key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFSearchFile", "BRIEFSearchFile", "Emulates BRIEF F5 key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFSearchNext", "BRIEFSearchNext", "Emulates BRIEF SHIFT-F5 key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFToggleColumnSelect", "BRIEFToggleColumnSelect", "Emulates BRIEF ALT-C key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFToggleLineSelect", "BRIEFToggleLineSelect", "Emulates BRIEF ALT-L key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFAltA", "BRIEFAltA", "Emulates BRIEF ALT-A key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFArrowLeft", "BRIEFArrowLeft", "Emulates BRIEF ARROW-LEFT key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFArrowRight", "BRIEFArrowRight", "Emulates BRIEF ARROW-RIGHT key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFCtrlArrowLeft", "BRIEFCtrlArrowLeft", "Emulates BRIEF Ctrl-ARROW-LEFT key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFCtrlArrowRight", "BRIEFCtrlArrowRight", "Emulates BRIEF Ctrl-ARROW-RIGHT key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFArrowUp", "BRIEFArrowUp", "Emulates BRIEF ARROW-UP key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
            CommandObj = _applicationObject.Commands.AddNamedCommand(objAddIn, "BRIEFArrowDown", "BRIEFArrowDown", "Emulates BRIEF ARROW-DOWN key", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            'CommandObj.AddControl(objCommandBar)
        Else
            'If you are not using events, you may wish to remove some of these to increase performance.
            'EnvironmentEvents.DTEEvents = CType(_applicationObject.Events.DTEEvents, EnvDTE.DTEEvents)
            'EnvironmentEvents.DocumentEvents = CType(_applicationObject.Events.DocumentEvents(Nothing), EnvDTE.DocumentEvents)
            'EnvironmentEvents.WindowEvents = CType(_applicationObject.Events.WindowEvents(Nothing), EnvDTE.WindowEvents)
            'EnvironmentEvents.TaskListEvents = CType(_applicationObject.Events.TaskListEvents(""), EnvDTE.TaskListEvents)
            'EnvironmentEvents.FindEvents = CType(_applicationObject.Events.FindEvents, EnvDTE.FindEvents)
            'EnvironmentEvents.OutputWindowEvents = CType(_applicationObject.Events.OutputWindowEvents(""), EnvDTE.OutputWindowEvents)
            'EnvironmentEvents.SelectionEvents = CType(_applicationObject.Events.SelectionEvents, EnvDTE.SelectionEvents)
            'EnvironmentEvents.SolutionItemsEvents = CType(_applicationObject.Events.SolutionItemsEvents, EnvDTE.ProjectItemsEvents)
            'EnvironmentEvents.MiscFilesEvents = CType(_applicationObject.Events.MiscFilesEvents, EnvDTE.ProjectItemsEvents)
            'EnvironmentEvents.DebuggerEvents = CType(_applicationObject.Events.DebuggerEvents, EnvDTE.DebuggerEvents)
        End If
    End Sub

    '''<summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
    '''<param name='disconnectMode'>Describes how the Add-in is being unloaded.</param>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnDisconnection(ByVal disconnectMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
    End Sub

    '''<summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification that the collection of Add-ins has changed.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
    End Sub

    '''<summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
    End Sub

    '''<summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
    End Sub

    '''<summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
    '''<param name='commandName'>The name of the command to determine state for.</param>
    '''<param name='neededText'>Text that is needed for the command.</param>
    '''<param name='status'>The state of the command in the user interface.</param>
    '''<param name='commandText'>Text requested by the neededText parameter.</param>
    '''<remarks></remarks>
    Public Sub QueryStatus(ByVal commandName As String, ByVal neededText As vsCommandStatusTextWanted, ByRef status As vsCommandStatus, ByRef commandText As Object) Implements IDTCommandTarget.QueryStatus
        status = vsCommandStatus.vsCommandStatusUnsupported
        If neededText = EnvDTE.vsCommandStatusTextWanted.vsCommandStatusTextWantedNone Then
            If commandName = "BRIEF.Connect.BRIEFHomeKey" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFEndKey" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFCopy" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFCut" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFPaste" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFLineDelete" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFDelete" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFPageUp" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFPageDown" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFUndo" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFToggleLineSelect" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFSearchFile" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFSearchNext" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFToggleColumnSelect" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFAltA" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowLeft" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowRight" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFCtrlArrowLeft" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFCtrlArrowRight" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowUp" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowDown" Then
                status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If
            'If commandName = "BRIEF.Connect.AddWSConstructor" Then
            'status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            'End If

        End If
    End Sub

    '''<summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
    '''<param name='commandName'>The name of the command to execute.</param>
    '''<param name='executeOption'>Describes how the command should be run.</param>
    '''<param name='varIn'>Parameters passed from the caller to the command handler.</param>
    '''<param name='varOut'>Parameters passed from the command handler to the caller.</param>
    '''<param name='handled'>Informs the caller if the command was handled or not.</param>
    '''<remarks></remarks>
    Public Sub Exec(ByVal commandName As String, ByVal executeOption As vsCommandExecOption, ByRef varIn As Object, ByRef varOut As Object, ByRef handled As Boolean) Implements IDTCommandTarget.Exec
        handled = False
        If (executeOption = vsCommandExecOption.vsCommandExecOptionDoDefault) Then
            If commandName = "BRIEF.Connect.BRIEFHomeKey" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFHomeKey()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFEndKey" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFEndKey()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFCopy" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFCopy()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFCut" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFCut()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFPaste" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFPaste()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFLineDelete" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFLineDelete()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFDelete" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFDelete()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFPageUp" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFPageUp()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFPageDown" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFPageDown()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFUndo" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFUndo()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFSearchFile" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFSearchFile()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFSearchNext" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFSearchNext()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFToggleColumnSelect" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFToggleColumnSelect()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFToggleLineSelect" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFToggleLineSelect()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFAltA" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFAltA()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowLeft" Then
                Keys.DTE = _applicationObject
                Keys.BRIEFArrowLeft(handled)
                If handled = False Then
                    _app.ExecuteCommand("Edit.CharLeft")
                End If
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowRight" Then
                Keys.DTE = _applicationObject
                Keys.BRIEFArrowRight(handled)
                If handled = False Then
                    _app.ExecuteCommand("Edit.CharRight")
                End If
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFCtrlArrowLeft" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFCtrlArrowLeft()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFCtrlArrowRight" Then
                handled = True
                Keys.DTE = _applicationObject
                Keys.BRIEFCtrlArrowRight()
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowUp" Then
                Keys.DTE = _applicationObject
                Keys.BRIEFArrowUp(handled)
                If handled = False Then
                    _app.ExecuteCommand("Edit.LineUp")
                End If
                Exit Sub
            End If
            If commandName = "BRIEF.Connect.BRIEFArrowDown" Then
                Keys.DTE = _applicationObject
                Keys.BRIEFArrowDown(handled)
                If handled = False Then
                    _app.ExecuteCommand("Edit.LineDown")
                End If
            End If
            Exit Sub
        End If
    End Sub

End Class
