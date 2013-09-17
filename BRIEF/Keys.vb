Option Explicit On
Option Strict Off
Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE90a
Imports EnvDTE100
Imports System.Diagnostics

Public Module Keys
    Public DTE As EnvDTE80.DTE2

    ' After installing this AddIn you'll find the following functions in Options->Keyboard
    ' prefixed with BRIEF.Connect.
    '
    ' BRIEFPaste      INSERT
    ' BRIEFCut        NUM-MINUS
    ' BRIEFCopy       NUM-PLUS
    ' BRIEFHomeKey    HOME
    ' BRIEFEndKey     END
    ' BRIEFUndo       NUM-ASTERISK
    ' BRIEFLineDelete ALT-D
    ' 
    ' BRIEFSearchFile F5
    ' BRIEFSearchNext SHIFT-F5

    ' for ALT-C and/or ALT-L to work, all of these keys must be mapped to the AddIn functions below
    '
    ' BRIEFToggleColumnSelect ALT-C
    ' BRIEFToggleLineSelect ALT-L
    ' BRIEFPageUp     PAGE UP
    ' BRIEFPageDown   PAGE DOWN
    ' BRIEFArrowDown  ARROW DOWN
    ' BRIEFArrowUp    ARROW UP
    ' BRIEFArrowLeft  ARROW LEFT
    ' BRIEFArrowRight ARROW RIGHT

    ' EditMakeUppercase CTRL-UP
    ' EditMakeLowercase CTRL-DOWN

    ' this is the interval, in seconds, between presses of HOME or END, to invoke 
    ' HOME-HOME-HOME or END-END-END functionality
    '
    ' 0.65 was too long
    '
    Private Const SECS_BETWEEN_HOME_END_KEYPRESS As Integer = 0.65
    Private ColumnSelectActive As Boolean
    Private LineSelectActive As Boolean
    Private AltAActive As Boolean
    Private SearchString As String

    Private Enum ClipModeEnum
        StreamMode = 0
        BoxMode = 1
        LineMode = 2
    End Enum

    Private Enum HomeMovementLevelEnum
        Text = 0
        Line = 1
        Page = 2
        Doc = 3
    End Enum

    Private Enum EndMovementLevelEnum
        Line = 1
        Page = 2
        Doc = 3
    End Enum

    Private NextHomeMovement As HomeMovementLevelEnum
    Private NextEndMovement As EndMovementLevelEnum
    Private LastEndPress As Date
    Private LastHomePress As Date
    Private NumLinesCopied As Integer
    Private ClipMode As ClipModeEnum
    Private Const ONE_SECOND_OA As Double = 0.00001157407407407407

    Private Function IsTimeElapsed(ByVal StartTime As Date, ByVal Seconds As Double) As Boolean
        Dim Duration As Double = (ONE_SECOND_OA * Seconds)
        Dim Elapsed As Double = (CDbl(Now.ToOADate()) - CDbl(StartTime.ToOADate()))
        Return CBool(Elapsed > Duration)
    End Function


    Public Sub BRIEFHomeKey()
        ' count as next press in sequence if 1 second or 
        ' less has elapsed since last press
        '
        ''If (DateDiff(DateInterval.Second, LastHomePress, Now()) <= SECS_BETWEEN_HOME_END_KEYPRESS) Then
        If Not IsTimeElapsed(LastHomePress, SECS_BETWEEN_HOME_END_KEYPRESS) Then
            NextHomeMovement += 1
        Else
            NextHomeMovement = HomeMovementLevelEnum.Text
        End If

        LastHomePress = Now()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        Select Case NextHomeMovement
            Case HomeMovementLevelEnum.Text
                sel.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText, ColumnSelectActive Or AltAActive)

            Case HomeMovementLevelEnum.Line
                sel.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn, ColumnSelectActive Or AltAActive)

            Case HomeMovementLevelEnum.Page
                ' need to figure out first visible line
                ''sel.ActivePoint.TryToShow(vsPaneShowHow.vsPaneShowTop)
                sel.LineUp(ColumnSelectActive Or LineSelectActive Or AltAActive, sel.ActivePoint.Line - sel.TextPane.StartPoint.Line)

            Case HomeMovementLevelEnum.Doc
                sel.StartOfDocument(ColumnSelectActive Or LineSelectActive Or AltAActive)
                NextHomeMovement = HomeMovementLevelEnum.Text

            Case Else
                NextHomeMovement = HomeMovementLevelEnum.Text
        End Select
    End Sub

    Public Sub ResetHomeAndEnd()
        NextEndMovement = EndMovementLevelEnum.Line
        NextHomeMovement = HomeMovementLevelEnum.Text
        LastEndPress = Now().AddSeconds(-2)
        LastHomePress = Now().AddSeconds(-2)
    End Sub

    Public Sub BRIEFEndKey()
        ' count as next press in sequence if 1 second or 
        ' less has elapsed since last press
        '
        'If (DateDiff(DateInterval.Second, LastEndPress, Now()) <= SECS_BETWEEN_HOME_END_KEYPRESS) Then
        If Not IsTimeElapsed(LastEndPress, SECS_BETWEEN_HOME_END_KEYPRESS) Then
            NextEndMovement += 1
        Else
            NextEndMovement = EndMovementLevelEnum.Line
        End If
        LastEndPress = Now()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        Select Case NextEndMovement
            Case EndMovementLevelEnum.Line
                sel.EndOfLine(ColumnSelectActive Or AltAActive)

            Case EndMovementLevelEnum.Page
                ' need to figure out lastst visible line
                Dim LastVisLine As Integer = sel.TextPane.StartPoint.Line + sel.TextPane.Height
                sel.LineDown(ColumnSelectActive Or LineSelectActive Or AltAActive, LastVisLine - sel.ActivePoint.Line)
                If Not ColumnSelectActive Or AltAActive Then sel.StartOfLine()

            Case EndMovementLevelEnum.Doc
                sel.EndOfDocument(ColumnSelectActive Or LineSelectActive Or AltAActive)
                NextEndMovement = EndMovementLevelEnum.Line

            Case Else
                NextEndMovement = EndMovementLevelEnum.Line
        End Select
    End Sub

    '' Map to keypad-plus
    Public Sub BRIEFCopy()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection
        If (sel.IsEmpty = True) Then
            Dim CharOffset As Integer = sel.ActivePoint.VirtualCharOffset
            sel.SelectLine()
            sel.Copy()
            sel.MoveToDisplayColumn(sel.ActivePoint.Line - 1, CharOffset)
            NumLinesCopied = 1
            ClipMode = ClipModeEnum.LineMode
        Else
            ClipMode = Math.Abs(CInt(sel.Mode = vsSelectionMode.vsSelectionModeBox))
            NumLinesCopied = (sel.BottomLine - sel.TopLine) + 1
            sel.Copy()
        End If

        ColumnSelectActive = False
        LineSelectActive = False
        AltAActive = False

        sel.Collapse()
    End Sub

    '' Map to keypad-minus
    Public Sub BRIEFCut()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        If (sel.IsEmpty = True) Then
            Dim CharOffset As Integer = sel.ActivePoint.VirtualDisplayColumn
            sel.SelectLine()
            sel.Cut()
            sel.MoveToDisplayColumn(sel.ActivePoint.Line, CharOffset)
            NumLinesCopied = 1
            ClipMode = ClipModeEnum.LineMode
        Else
            ClipMode = Math.Abs(CInt(sel.Mode = vsSelectionMode.vsSelectionModeBox))
            NumLinesCopied = (sel.BottomLine - sel.TopLine) + 1
            sel.Cut()
        End If

        LineSelectActive = False
        ColumnSelectActive = False
        AltAActive = False

        Return
    End Sub

    '' Map to INS
    Public Sub BRIEFPaste()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        Dim CharOffset As Integer = sel.ActivePoint.VirtualDisplayColumn
        If ClipMode = ClipModeEnum.LineMode Then
            sel.StartOfLine()
            sel.Paste()
            sel.MoveToDisplayColumn(sel.ActivePoint.Line, CharOffset)
            Return
        End If

        ' If we paste while the cursor is on a blank line while in box mode with a selection
        ' more than one line high, it inserts the selection as lines not columns.  So the 
        ' sleazy work-around is to make the current line not blank, insert a period, select 
        ' it and then paste (which gets rid of the period.)
        '
        If sel.ActivePoint.LineLength = 0 Then
            sel.Insert(".")
            sel.CharLeft(True)
        End If

        sel.Paste()
        Try
            If ClipMode = ClipModeEnum.BoxMode Then sel.MoveToDisplayColumn(sel.ActivePoint.Line + NumLinesCopied, CharOffset)
        Catch ex As Exception
        End Try

    End Sub

    '' Map to Alt-L
    Public Sub BRIEFToggleLineSelect()
        ResetHomeAndEnd()
        If AltAActive Or ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Collapse()
            AltAActive = False
            ColumnSelectActive = False
        End If
        Dim sel As TextSelection = DTE.ActiveDocument.Selection
        LineSelectActive = Not LineSelectActive
        If LineSelectActive = False Then
            DTE.ActiveDocument.Selection.Collapse()
        Else
            sel.SelectLine()
        End If
    End Sub

    ' Map to ALT-C
    Public Sub BRIEFToggleColumnSelect()
        ResetHomeAndEnd()
        If LineSelectActive Or AltAActive Then
            DTE.ActiveDocument.Selection.Collapse()
            AltAActive = False
            LineSelectActive = False
        End If
        ColumnSelectActive = Not ColumnSelectActive
        LineSelectActive = False
        AltAActive = False
        If ColumnSelectActive = False Then DTE.ActiveDocument.Selection.Collapse()
    End Sub

    '' Map to Alt-A
    Public Sub BRIEFAltA()
        ResetHomeAndEnd()
        If LineSelectActive Or ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Collapse()
            LineSelectActive = False
            ColumnSelectActive = False
        End If

        Dim sel As TextSelection = DTE.ActiveDocument.Selection
        AltAActive = Not AltAActive
        If AltAActive = False Then
            DTE.ActiveDocument.Selection.Collapse()
        End If
    End Sub

    '' Map to Alt-D
    Public Sub BRIEFLineDelete()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        Dim CharOffset As Integer = sel.ActivePoint.VirtualDisplayColumn
        sel.SelectLine()
        sel.Delete()
        sel.MoveToDisplayColumn(sel.ActivePoint.Line, CharOffset)

        ColumnSelectActive = False
        LineSelectActive = False
        AltAActive = False

        Return
    End Sub

    '' Map to del key
    Public Sub BRIEFDelete()
        ResetHomeAndEnd()
        Dim sel As TextSelection = DTE.ActiveDocument.Selection
        sel.Delete()
        ColumnSelectActive = False
        LineSelectActive = False
        AltAActive = False
        Return
    End Sub

    '' Map to keypad-asterisk
    Public Sub BRIEFUndo()
        ResetHomeAndEnd()

        DTE.ActiveDocument.Undo()
    End Sub

    '' Map to F5
    Public Sub BRIEFSearchFile()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        If (Not sel.IsEmpty) And (sel.TopPoint.Line = sel.BottomPoint.Line) Then
            SearchString = sel.Text
        End If


        'If (Not sel.IsEmpty) And (sel.Mode <> vsSelectionMode.vsSelectionModeBox) Then
        '    SearchString = InputBox("Text to find:", "BRIEF Emulation: Find in selection", SearchString)
        '    If (SearchString <> "") Then
        '        Dim pos = InStr(sel.Text, SearchString)
        '        Dim TopLine = sel.TopLine
        '        If (pos > 0) Then
        '            sel.MoveToAbsoluteOffset(sel.TopPoint.AbsoluteCharOffset + pos)
        '            sel = DTE.ActiveDocument.Selection
        '            sel.MoveToAbsoluteOffset(sel.ActivePoint.AbsoluteCharOffset - (sel.TopLine - TopLine))
        '        End If
        '    End If
        'Else

        SearchString = InputBox("Text to find:", "BRIEF Emulation: Find in file", SearchString)
        If (SearchString <> "") Then
            sel.FindText(SearchString)
        End If

        'End If
    End Sub

    '' Map to shift-F5
    Public Sub BRIEFSearchNext()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        'If (Not sel.IsEmpty) And (sel.Mode <> vsSelectionMode.vsSelectionModeBox) Then
        '    If (SearchString <> "") Then
        '        Dim pos = InStr(sel.Text, SearchString)
        '        Dim TopLine = sel.TopLine
        '        If (pos > 0) Then
        '            sel.MoveToAbsoluteOffset(sel.TopPoint.AbsoluteCharOffset + pos)
        '            sel = DTE.ActiveDocument.Selection
        '            sel.MoveToAbsoluteOffset(sel.ActivePoint.AbsoluteCharOffset - (sel.TopLine - TopLine))
        '        End If
        '    End If
        'Else
        If (SearchString <> "") Then
            sel.FindText(SearchString)
        End If
        'End If
    End Sub

    '' beginnings of an advanced search dialog
    Private Sub TestCreateDialog()

        Dim f As New System.Windows.Forms.Form
        f.Height = 200
        f.Width = 200
        f.Text = "test"
        Dim b As New System.Windows.Forms.Button
        b.Height = 50
        b.Width = 75
        b.Text = "Ok"
        b.DialogResult = System.Windows.Forms.DialogResult.OK
        f.Controls.Add(b)
        b.Left = 100
        b.Top = 20
        f.AcceptButton = b
        f.TopMost = True
        Dim res As System.Windows.Forms.DialogResult = f.ShowDialog()
    End Sub

    ' Map to arrow left
    Public Sub BRIEFArrowLeft()
        ResetHomeAndEnd()

        If ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Mode = vsSelectionMode.vsSelectionModeBox
            DTE.ActiveDocument.Selection.CharLeft(True, 1)
        Else
            DTE.ActiveDocument.Selection.CharLeft(AltAActive, 1)
        End If
    End Sub

    ' Map to arrow right
    Public Sub BRIEFArrowRight()
        ResetHomeAndEnd()

        If ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Mode = vsSelectionMode.vsSelectionModeBox
            DTE.ActiveDocument.Selection.CharLeft(True, -1)
        Else
            DTE.ActiveDocument.Selection.CharLeft(AltAActive, -1)
        End If
    End Sub

    ' Map to arrow left
    Public Sub BRIEFCtrlArrowLeft()
        ResetHomeAndEnd()

        If ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Mode = vsSelectionMode.vsSelectionModeBox
            DTE.ActiveDocument.Selection.WordLeft(True, 1)
        Else
            DTE.ActiveDocument.Selection.WordLeft(AltAActive, 1)
        End If
    End Sub

    ' Map to arrow right
    Public Sub BRIEFCtrlArrowRight()
        ResetHomeAndEnd()

        If ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Mode = vsSelectionMode.vsSelectionModeBox
            DTE.ActiveDocument.Selection.WordLeft(True, -1)
        Else
            DTE.ActiveDocument.Selection.WordLeft(AltAActive, -1)
        End If
    End Sub



    ' Map to arrow up
    Public Sub BRIEFArrowUp()
        ResetHomeAndEnd()

        If ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Mode = vsSelectionMode.vsSelectionModeBox
        End If
        DTE.ActiveDocument.Selection.LineDown(ColumnSelectActive Or LineSelectActive Or AltAActive, -1)

    End Sub

    ' Map to arrow down
    Public Sub BRIEFArrowDown()
        ResetHomeAndEnd()

        If ColumnSelectActive Then
            DTE.ActiveDocument.Selection.Mode = vsSelectionMode.vsSelectionModeBox
        End If
        DTE.ActiveDocument.Selection.LineDown(ColumnSelectActive Or LineSelectActive Or AltAActive, 1)
    End Sub

    '' Map to page up
    Public Sub BRIEFPageUp()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        If (Not sel.IsEmpty) And (sel.Mode = vsSelectionMode.vsSelectionModeBox) Then
            sel.PageUp(True, 1)
        Else
            sel.PageUp(LineSelectActive Or AltAActive, 1)
        End If
    End Sub

    '' Map to page down
    Public Sub BRIEFPageDown()
        ResetHomeAndEnd()

        Dim sel As TextSelection = DTE.ActiveDocument.Selection

        If (Not sel.IsEmpty) And (sel.Mode = vsSelectionMode.vsSelectionModeBox) Then
            sel.PageDown(True, 1)
        Else
            sel.PageDown(LineSelectActive Or AltAActive, 1)
        End If
    End Sub
End Module
