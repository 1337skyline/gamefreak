VERSION 5.00
Begin VB.Form MainMenu 
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Game Stations"
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tech Schedule"
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Coupons"
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gamers"
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   480
      X2              =   2520
      Y1              =   7080
      Y2              =   7080
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Gamers.Hide
Coupons.Hide
TechShed.Hide
If GameStations.Visible = False Then
GameStations.Show
GameStations.Image1.Visible = True
Else
GameStations.Image1.Visible = False
GameStations.Hide
End If
End Sub

Private Sub Command2_Click()
GameStations.Hide
Coupons.Hide
TechShed.Hide
If Gamers.Visible = False Then
Gamers.Show
Gamers.Image1.Visible = True
Else
Gamers.Image1.Visible = False
Gamers.Hide
End If
End Sub

Private Sub Command3_Click()
GameStations.Hide
Gamers.Hide
TechShed.Hide
If Coupons.Visible = False Then
Coupons.Show
Coupons.Image1.Visible = True
Else
Coupons.Image1.Visible = False
Coupons.Hide
End If
End Sub

Private Sub Command4_Click()
GameStations.Hide
Gamers.Hide
Coupons.Hide
If TechShed.Visible = False Then
TechShed.Show
TechShed.Image1.Visible = True
Else
TechShed.Image1.Visible = False
TechShed.Hide
End If
End Sub

