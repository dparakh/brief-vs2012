Option Strict Off
Option Explicit Off
Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE90a
Imports EnvDTE100
Imports System.Diagnostics

Public Module RecordingModule
	public Dim DTE as EnvDTE80.DTE2


    Sub AddWSConstructor()
        DTE.ActiveDocument.Selection.StartOfDocument()
        DTE.ExecuteCommand("Edit.Find")
        DTE.Find.FindWhat = "public sub new"
        DTE.Find.Target = vsFindTarget.vsFindTargetCurrentDocument
        DTE.Find.MatchCase = False
        DTE.Find.MatchWholeWord = False
        DTE.Find.Backwards = False
        DTE.Find.MatchInHiddenText = False
        DTE.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxLiteral
        DTE.Find.Action = vsFindAction.vsFindActionFind
        If (DTE.Find.Execute() = vsFindResult.vsFindResultNotFound) Then
            Throw New System.Exception("vsFindResultNotFound")
        End If
        DTE.Find.FindWhat = "End Sub"
        If (DTE.Find.Execute() = vsFindResult.vsFindResultNotFound) Then
            Throw New System.Exception("vsFindResultNotFound")
        End If
        DTE.Windows.Item("{CF2DDC32-8CAD-11D2-9302-005345000000}").Close()
        DTE.ActiveDocument.Selection.LineDown()
        DTE.ActiveDocument.Selection.NewLine()
        DTE.ActiveDocument.Selection.Text = "   Public Sub New(ServerUrl As String)"
        DTE.ActiveDocument.Selection.NewLine()
        DTE.ActiveDocument.Selection.Text = "       MyBase.New()"
        DTE.ActiveDocument.Selection.NewLine()
        DTE.ActiveDocument.Selection.Text = "       Me.Url = ServerUrl + ""/OutlookSync/OLSWebService.asmx"""
        DTE.ActiveDocument.Selection.NewLine()
        DTE.ActiveDocument.Selection.Text = "   End Sub"
        DTE.ActiveDocument.Selection.NewLine()

    End Sub
End Module
