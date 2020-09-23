VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Scrolling Label by PRzEm..."
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ramAbout 
      Height          =   3135
      Left            =   740
      TabIndex        =   1
      Top             =   195
      Width           =   4680
      Begin VB.Label lblN_Email 
         AutoSize        =   -1  'True
         Caption         =   "E-mail me at:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmAbout.frx":000C
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "admin@videoclips.prv.pl"
         DragIcon        =   "frmAbout.frx":0170
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   1725
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         Caption         =   "PRzEm"
         Height          =   195
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblN_Author 
         AutoSize        =   -1  'True
         Caption         =   "Author:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Scrolling Label ver: 1.0"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblN_Name 
         AutoSize        =   -1  'True
         Caption         =   "UserControl:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "frmAbout.frx":047A
      Stretch         =   -1  'True
      Top             =   200
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdOK_Click()
Unload frmAbout
Set frmAbout = Nothing
End Sub
Private Sub lblEmail_DragDrop(Source As Control, X As Single, Y As Single)
If Source Is lblEmail Then
 With lblEmail
  .Font.Underline = False
  .ForeColor = vbBlack
  Call ShellExecute(0&, vbNullString, "mailto:" & .Caption, vbNullString, vbNullString, vbNormalFocus)
 End With
End If

End Sub
Private Sub lblEmail_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
If State = vbLeave Then
 With lblEmail
  .Drag vbEndDrag
  .Font.Underline = False
  .ForeColor = vbBlack
 End With
End If

End Sub
Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lblEmail
 .ForeColor = vbBlue
 .Font.Underline = True
 .Drag vbBeginDrag
End With

End Sub
