VERSION 5.00
Begin VB.Form GameStations 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   2595
   ClientTop       =   0
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GameStations.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Gamers"
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   2775
      Left            =   600
      TabIndex        =   32
      Top             =   5280
      Width           =   2415
      Begin VB.TextBox Text6 
         DataField       =   "GFCoinsLeft"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   720
         TabIndex        =   42
         Text            =   "Text6"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text5 
         DataField       =   "TimeUsed"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text4 
         DataField       =   "GFCoinsUsed"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   360
         TabIndex        =   40
         Text            =   "Text4"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   38
         Text            =   "Combo1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1200
         TabIndex        =   36
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Play!"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   2040
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   1560
         Max             =   1
         Min             =   24
         TabIndex        =   34
         Top             =   960
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "1"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pick A Gamer:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time playing:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lets play"
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   375
      Left            =   720
      TabIndex        =   30
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check2"
      DataField       =   "Occupied"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      DataField       =   "Enabled"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   1800
      TabIndex        =   25
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text2 
      DataField       =   "Earnings"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "TimeUsed"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   3480
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Stations"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GameStations"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   3000
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   3000
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblview 
      BackColor       =   &H8000000E&
      Caption         =   "Station #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   29
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblview 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   555
      Index           =   0
      Left            =   1440
      TabIndex        =   28
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Earnings"
      Height          =   255
      Left            =   840
      TabIndex        =   27
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Occuipied"
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Used"
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   20
      Left            =   10800
      TabIndex        =   0
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             19"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   19
      Left            =   9000
      TabIndex        =   1
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   18
      Left            =   7200
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             17"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   17
      Left            =   5400
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             16"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   16
      Left            =   3600
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   15
      Left            =   10800
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   14
      Left            =   9000
      TabIndex        =   6
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   13
      Left            =   7200
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   12
      Left            =   5400
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   11
      Left            =   3600
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "             10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   10
      Left            =   10800
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   9
      Left            =   9000
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   8
      Left            =   7200
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   7
      Left            =   5400
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   6
      Left            =   3600
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   5
      Left            =   10800
      TabIndex        =   15
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   4
      Left            =   9000
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   3
      Left            =   7200
      TabIndex        =   17
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   2
      Left            =   5400
      TabIndex        =   18
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H8000000E&
      Caption         =   "              1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Index           =   1
      Left            =   3600
      TabIndex        =   19
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   630
      Left            =   -2588
      Picture         =   "GameStations.frx":E5C5
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   2790
   End
End
Attribute VB_Name = "GameStations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Data2.Refresh
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Clear form" Then
VScroll1.Value = 1
Else
Frame1.Enabled = True
Command2.Caption = "Clear form"
End If
End Sub

Private Sub Command3_Click()
If (Text6 >= (VScroll1.Value * 15)) Then
Data2.Recordset.Fields ["TimeUsed"] = Text5 - VScroll1.Value
Data2.Recordset.Fields ["GFCoinsLeft"] = Text6 - (VScroll1.Value * 15)
Data2.Recordset.Fields ["GFCoinsUsed"] = Text4 + (VScroll1.Value * 15)
VScroll1.Value = 1
Command2.Caption = "Lets play"
Frame1.Enabled = False
End If
End Sub

Private Sub Command4_Click()
VScroll1.Value = 1
Command2.Caption = "Lets play"
Frame1.Enabled = False
End Sub

Private Sub Form_Activate()
Data2.Recordset.MoveLast
Data2.Recordset.MoveFirst
Do While Data2.Recordset.EOF = False
Combo1.AddItem Data2.Recordset(0)
Data2.Recordset.MoveNext
Loop
For i = 1 To 20
Next
End Sub


Private Sub lblSquare_Click(Index As Integer)
If Index > 0 Then
    lblview(0).Caption = Index
End If

Data1.RecordSource = "SELECT * FROM GameStations WHERE Station=" & Index + 1
Data1.Refresh
End Sub

Private Sub VScroll1_Change()
Text3.Text = VScroll1.Value
End Sub
