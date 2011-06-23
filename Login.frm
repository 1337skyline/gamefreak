VERSION 5.00
Begin VB.Form Login1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":0000
   ScaleHeight     =   5388.397
   ScaleMode       =   0  'User
   ScaleWidth      =   2802.753
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get In!"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Login.frx":40E1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Height          =   390
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   780
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label LoginProcess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please enter Username and Password."
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   6120
      Width           =   855
   End
End
Attribute VB_Name = "Login1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtUserName = "SkullCrusherz" And txtPassword = "OMGWTFBBQ" Then
        LoginSucceeded = True
        LoginProcess.Caption = "Success"
        LoginProcess.ForeColor = &HC000&
        Me.Hide
        MainMenu.Show
    Else
        LoginProcess.Caption = "Incorrect UserName/Paaword, Please try again."
        LoginProcess.ForeColor = &HFF&
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

