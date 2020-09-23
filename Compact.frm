VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Compact Database"
   ClientHeight    =   1020
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2910
   Icon            =   "Compact.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Repair"
      Height          =   465
      Left            =   450
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin VB.Menu to 
      Caption         =   "Tools"
      Begin VB.Menu CC 
         Caption         =   "Clean Compact"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub CC_Click()
CC.Checked = Not CC.Checked
End Sub

Private Sub Command1_Click()

On Error GoTo MSGER
Dim OFName As OPENFILENAME
OFName.lStructSize = Len(OFName)
'Set the parent window
OFName.hwndOwner = Me.hWnd
'Set the application's instance
OFName.hInstance = App.hInstance
'Select a filter
OFName.lpstrFilter = "Your Database (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
'create a buffer for the file
OFName.lpstrFile = Space$(254)
'set the maximum length of a returned file
OFName.nMaxFile = 255
'Create a buffer for the file title
OFName.lpstrFileTitle = Space$(254)
'Set the maximum length of a returned file title
OFName.nMaxFileTitle = 255
'Set the initial directory
OFName.lpstrInitialDir = App.Path
'Set the title
OFName.lpstrTitle = "Please Select your Database - VBSolutions2001.com"
'No flags
OFName.flags = 0

'Show the 'Open File'-dialog

If GetOpenFileName(OFName) Then
    dbsfle = Trim$(OFName.lpstrFile)
Else
    Exit Sub
End If

nme = Mid(dbsfle, InStrRev(dbsfle, "\") + 1)
nme = Left(nme, Len(nme) - 4)

If Dir(dbsfle) = "" Then MsgBox "Error. Contact Tech."

If Dir(App.Path & "\" & nme & ".CPT") = "" Then
Else
    Kill App.Path & "\" & nme & ".CPT"
End If

restart:

If Len(pass) > 5 Then
    DBEngine.CompactDatabase dbsfle, App.Path & "\" & nme & ".CPT", , , pass
Else
    DBEngine.CompactDatabase dbsfle, App.Path & "\" & nme & ".CPT"
End If

If Dir(App.Path & "\" & nme & ".OLD") = "" Then
Else
    Kill App.Path & "\" & nme & ".OLD"
End If

Name dbsfle As App.Path & "\" & nme & ".OLD"

If Len(pass) > 5 Then

    DBEngine.CompactDatabase App.Path & "\" & nme & ".CPT", dbsfle, , , pass
Else
    DBEngine.CompactDatabase App.Path & "\" & nme & ".CPT", dbsfle
End If

MsgBox "Finished Compacting " & dbsfle

If CC.Checked Then
    Kill App.Path & "\" & nme & ".CPT"
    Kill App.Path & "\" & nme & ".OLD"
End If

Exit Sub
MSGER:

If Err.Number = 3031 Then

    If Len(pass) > 0 Then answer = MsgBox(Right(pass, Len(pass) - 5) & " is a Invalid Password.", vbCritical + vbOKCancel, "Invalid Password")
    If answer = vbCancel Then Exit Sub
    pass = ";pwd=" & InputBox("Enter Password.", "VBSolutions2001.com")
 
    Resume restart

End If

answer = MsgBox(Err.Number & " In " & Name & " Line Number: " & Erl & " - " & Err.Description & " Click on Yes to RESUME, on No to RESUME NEXT, And On Cancel to EXIT SUB. Do You Want To RESUME?", vbYesNoCancel + vbCritical, "Error Message")

If answer = vbYes Then Resume

If answer = vbNo Then Resume Next
End Sub
