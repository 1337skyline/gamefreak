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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Update"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New order"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Coupons.frx":E5C5
      Height          =   5535
      Left            =   2520
      TabIndex        =   0
      Top             =   2760
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Coupon Allowed:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Freak level:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Times used:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   2160
      Width           =   975
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
