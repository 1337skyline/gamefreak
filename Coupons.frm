VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Coupons 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   2595
   ClientTop       =   0
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Coupons.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Done!"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   4815
      Begin VB.CommandButton Command3 
         Caption         =   "Forword"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         DataField       =   "Flevel"
         DataSource      =   "Data1"
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text3 
         DataField       =   "TimesUsed"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         DataField       =   "TimeCoupon"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Freak level:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Times used:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time Coupon Allowed:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Flevel"
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Update"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New level"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Coupons.frx":E5C5
      Height          =   4215
      Left            =   2520
      TabIndex        =   0
      Top             =   3120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   630
      Left            =   -2588
      Picture         =   "Coupons.frx":E5D9
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   2790
   End
End
Attribute VB_Name = "Coupons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If i = 1 Then
Data1.Recordset.Update
Else
If i = 2 Then
Data1.Recordset.Update
Else
Data1.Recordset.Delete
End If
End If
Frame1.Enabled = False
Command1.Enabled = False
Command8.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command1.Visible = False
Command8.Visible = False
End Sub

Private Sub Command2_Click()
If Data1.EOFAction = False Then

Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command3_Click()
Data1.Recordset.MoveNext
End Sub

Private Sub Command5_Click()
Frame1.Enabled = True
Command1.Enabled = True
Command8.Enabled = True
Command1.Visible = True
Command8.Visible = True
Data1.Recordset.AddNew
i = 1
End Sub

Private Sub Command6_Click()
Frame1.Enabled = True
Command1.Enabled = True
Command8.Enabled = True
Command1.Visible = True
Command8.Visible = True
Command2.Enabled = True
Command3.Enabled = True

i = 2
End Sub

Private Sub Command7_Click()
Frame1.Enabled = True
Command1.Enabled = True
Command8.Enabled = True
Command1.Visible = True
Command8.Visible = True
Command2.Enabled = True
Command3.Enabled = True

i = 3
End Sub

Private Sub Command8_Click()
Frame1.Enabled = False
Command1.Enabled = False
Command8.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command1.Visible = False
Command8.Visible = False
End Sub

Private Sub Form_Load()

End Sub
