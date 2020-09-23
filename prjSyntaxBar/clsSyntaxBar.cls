VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSyntaxBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDTExtensibility

'Our VBE instance
Public gVBInst As VBIDE.VBE

'Make sure we can handle the command bar's events
Public WithEvents BtnHandler As CommandBarEvents
Attribute BtnHandler.VB_VarHelpID = -1

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    'Add-in is now connected
   
    'Set the VB instance
    Set gVBInst = VBInst
   
    'Make sure we are supposed to load now
    If ConnectMode = vbext_cm_Startup Then
    
        Dim cbMenu As Object
  
        On Error GoTo AddCommandBarErr
           
        'Load our command bar 'prjSyntaxBar'
        gVBInst.CommandBars.Add "prjSyntaxBar", msoBarTop
        
        'Make sure that in loaded properly
        Set cbMenu = gVBInst.CommandBars("prjSyntaxBar")
        
        If cbMenu Is Nothing Then
            'Could not add command bar for some reason
            MsgBox "Fatal Error (1) Adding Command Bar!"
            Exit Sub
        End If
        
        'Make our command bar visible
        cbMenu.Visible = True
    
        'Button which when pressed, will display some basic
        'information about the project
        Dim mcbCmdBarInfo As Office.CommandBarButton
        
        'Add button to command bar
        Set mcbCmdBarInfo = gVBInst.CommandBars("prjSyntaxBar").Controls.Add(1)
        
        'Set button caption, and style to show caption
        mcbCmdBarInfo.Caption = "View Project Information"
        mcbCmdBarInfo.Style = msoButtonCaption
        
        'Turn on event handling for the button
        Set Me.BtnHandler = gVBInst.Events.CommandBarEvents(mcbCmdBarInfo)
    
        On Error GoTo 0

    End If

    Exit Sub
AddCommandBarErr:
    MsgBox "Fatal Error (Other) Adding Command Bar!"
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    'Add-in is now disconnected
   
    'Remove the command bar
    gVBInst.CommandBars("prjSyntaxBar").Delete
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    ' Comment to prevent procedure from being deleted on compilation.
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    ' Comment to prevent procedure from being deleted on compilation.
End Sub

'Button is clicked
Private Sub BtnHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim sInfo As String
    sInfo = "Project Information: "
    sInfo = sInfo & vbCrLf & vbCrLf & "Name: " & gVBInst.ActiveVBProject.Name
    sInfo = sInfo & vbCrLf & vbCrLf & "etc..."
    MsgBox sInfo
End Sub

