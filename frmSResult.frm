VERSION 5.00
Begin VB.Form frmSearchResults 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search Results"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstSearch 
      Height          =   2205
      Left            =   0
      TabIndex        =   1
      Top             =   280
      Width           =   4935
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "Always on top"
      Height          =   255
      Left            =   40
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub chkOnTop_Click()

If chkOnTop.Value = 1 Then
               SetWindowPos frmSearchResults.hwnd, HWND_TOPMOST, frmSearchResults.Left / 15, _
                            frmSearchResults.Top / 15, frmSearchResults.Width / 15, _
                            frmSearchResults.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
               SetWindowPos frmSearchResults.hwnd, HWND_NOTOPMOST, frmSearchResults.Left / 15, _
                            frmSearchResults.Top / 15, frmSearchResults.Width / 15, _
                            frmSearchResults.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If

End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then Exit Sub    ' Window is minimized, exit

On Error GoTo ErrHandleHResize
Me.lstSearch.Height = Me.Height - 720

ErrHandleHResize:
On Error GoTo ErrHandleWResize
Me.lstSearch.Width = Me.Width - 125

ErrHandleWResize:
Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #FileNum
    If FileExists(sourcePath$ + "search.tmp") Then Kill sourcePath$ & "search.tmp"
    
End Sub

Private Sub lstSearch_Click()

' ## UNDER CONSTRUCTION ## '
    
    If lstSearch.ListCount = 0 Then Exit Sub
    
End Sub
