Attribute VB_Name = "modAddinList"
Option Explicit

Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

Sub Main()
    Dim rc As Long
    rc = WritePrivateProfileString("Add-Ins32", "prjSyntaxBar.clsSyntaxBar", "3", "VBADDIN.INI")
    MsgBox "VBADDIN.INI Modified Successfully"
End Sub

