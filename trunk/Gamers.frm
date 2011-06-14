VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Gamers 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   2595
   ClientTop       =   0
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Gamers.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Update"
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   960
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "GTID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit A Pro"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete A Noob"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add A Newbie"
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Log"
      Top             =   7920
      Width           =   2220
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "CouponsAllowed"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox Text1 
      DataField       =   "FLevel"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   5
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "FLevel"
      Top             =   1920
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Gamers"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      DataField       =   "TimeUsed"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   8
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Time Used"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "PhoneNumber"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Phone Number:"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "GFCoinsUsed"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   7
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "GF Coins Used:"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "LastName"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   3
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Last Name"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "GFCoinsLeft"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   6
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "GF Coins Left:"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   2
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "First Name"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "GamerTag"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Gamer Tag"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Gamers.frx":E5C5
      Left            =   960
      List            =   "Gamers.frx":E5C7
      TabIndex        =   0
      Text            =   "GamerTag ID#"
      Top             =   1320
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Gamers.frx":E5C9
      Height          =   3255
      Left            =   960
      TabIndex        =   10
      Top             =   5040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone Number:"
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Name:"
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name:"
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gamer Tag:"
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Freak Level:"
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Used"
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GF Coins Used:"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GF Coins Left:"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Coupons Allowed:"
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   630
      Left            =   -2588
      Picture         =   "Gamers.frx":E5DD
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   2790
   End
End
Attribute VB_Name = "Gamers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Option Explicit


Private Sub Combo1_Click()
Data1.Recordset.FindFirst "GTID=" & Combo1.Text
End Sub

Private Sub Command1_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
If Data1.Recordset.RecordCount = 0 Then
MsgBox "There are no gamers to edit.", vbCritical, "Error"
Exit Sub
End If
If Text1(1).Locked = False Then
For i = 1 To 8
Text1(i).Locked = True
Command5.Visible = False
Command4.Visible = False
Next i
Else
For i = 1 To 8
Text1(i).Locked = False
Command5.Visible = True
Command4.Visible = True
Command3.Visible = False
Next i
End If
End Sub

Private Sub Command5_Click()
Data1.Recordset.Update
Data1.Recordset.MoveLast
End Sub

Private Sub Form_Activate()
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF = False
Combo1.AddItem Data1.Recordset(0)
Data1.Recordset.MoveNext
Loop
Data1.Recordset.MoveFirst
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
End Select
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then Exit Sub
Data2.RecordSource = "SELECT Log.LogCode, Log.GTID, Log.CoupunUsed, Log.DateLog FROM Log INNER JOIN Gamers ON Log.GTID = Gamers.GTID WHERE (Log.GTID)=" & Text2
Data2.Refresh
End Sub
