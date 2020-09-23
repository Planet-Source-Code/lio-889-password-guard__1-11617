VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okay"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame frameAbout 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   25
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6975
   End
   Begin VB.Image imgNew 
      Height          =   150
      Left            =   4245
      Picture         =   "frmAbout.frx":0CCA
      Top             =   675
      Width           =   360
   End
   Begin VB.Image imgTitle 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0D39
      Top             =   240
      Width           =   435
   End
   Begin VB.Label lblThanx 
      AutoSize        =   -1  'True
      Caption         =   "Thanks for using my software :-D"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblEmailCap 
      AutoSize        =   -1  'True
      Caption         =   "Email address:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label lblHomePageCap 
      AutoSize        =   -1  'True
      Caption         =   "Home Page:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "lio_889@ziplip.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      MouseIcon       =   "frmAbout.frx":1903
      TabIndex        =   6
      Top             =   1440
      Width           =   1875
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      Caption         =   "http://www.geocities.com/lio889"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      MouseIcon       =   "frmAbout.frx":1C0D
      TabIndex        =   5
      Top             =   1680
      Width           =   3270
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "This computer program is 100% FREEWARE. You may copy and/or distribute this program in any way you may find it useful."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (C) 2000 Khaery Rida."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2835
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.1 (Freeware)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2370
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Password Guard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3960
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
frmAbout.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
On Error Resume Next
frmAbout.Hide

End Sub

Private Sub lblEmail_Click()
RetVal = ShellExecute(Me.hwnd, vbNullString, "mailto:lio_889@ziplip.com?subject=" & MainTitle, vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL)

End Sub

Private Sub lblHomePage_Click()
RetVal = ShellExecute(Me.hwnd, vbNullString, "http://www.geocities.com/lio889", vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL)

End Sub
