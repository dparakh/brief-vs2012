﻿Option Strict Off
Option Explicit Off
Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE90a
Imports EnvDTE100
Imports System.Diagnostics

Public Module EnvironmentEvents
    Public DTE As EnvDTE80.DTE2

#Region "Automatically generated code, do not modify"
    'Automatically generated code, do not modify
    'Event Sources Begin
    <System.ContextStaticAttribute()> Public WithEvents BuildEvents As EnvDTE.BuildEvents
    <System.ContextStaticAttribute()> Public WithEvents CodeModelEvents As EnvDTE80.CodeModelEvents
    <System.ContextStaticAttribute()> Public WithEvents DebuggerEvents As EnvDTE.DebuggerEvents
    <System.ContextStaticAttribute()> Public WithEvents DebuggerExpressionEvaluationEvents As EnvDTE80.DebuggerExpressionEvaluationEvents
    <System.ContextStaticAttribute()> Public WithEvents DebuggerProcessEvents As EnvDTE80.DebuggerProcessEvents
    <System.ContextStaticAttribute()> Public WithEvents DocumentEvents As EnvDTE.DocumentEvents
    <System.ContextStaticAttribute()> Public WithEvents DTEEvents As EnvDTE.DTEEvents
    <System.ContextStaticAttribute()> Public WithEvents FindEvents As EnvDTE.FindEvents
    <System.ContextStaticAttribute()> Public WithEvents MiscFilesEvents As EnvDTE.ProjectItemsEvents
    <System.ContextStaticAttribute()> Public WithEvents OutputWindowEvents As EnvDTE.OutputWindowEvents
    <System.ContextStaticAttribute()> Public WithEvents SolutionItemsEvents As EnvDTE.ProjectItemsEvents
    <System.ContextStaticAttribute()> Public WithEvents ProjectsEvents As EnvDTE.ProjectsEvents
    <System.ContextStaticAttribute()> Public WithEvents SelectionEvents As EnvDTE.SelectionEvents
    <System.ContextStaticAttribute()> Public WithEvents SolutionEvents As EnvDTE.SolutionEvents
    <System.ContextStaticAttribute()> Public WithEvents TaskListEvents As EnvDTE.TaskListEvents
    <System.ContextStaticAttribute()> Public WithEvents TextDocumentKeyPressEvents As EnvDTE80.TextDocumentKeyPressEvents
    <System.ContextStaticAttribute()> Public WithEvents TextEditorEvents As EnvDTE.TextEditorEvents
    <System.ContextStaticAttribute()> Public WithEvents WindowEvents As EnvDTE.WindowEvents
    'Event Sources End
    'End of automatically generated code
#End Region


    ' Private Sub TextEditorEvents_LineChanged(StartPoint As EnvDTE.TextPoint, EndPoint As EnvDTE.TextPoint, Hint As Integer) Handles TextEditorEvents.LineChanged

    'End Sub

End Module

