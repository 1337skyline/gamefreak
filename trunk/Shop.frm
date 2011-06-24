VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form TechShed 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   2595
   ClientTop       =   0
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Shop.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GameStations"
      Top             =   8400
      Width           =   1740
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Done"
      Height          =   375
      Left            =   9480
      TabIndex        =   28
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Done"
      Height          =   375
      Left            =   3240
      TabIndex        =   27
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   1695
      Left            =   6600
      TabIndex        =   14
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox Text10 
         DataField       =   "FirstName"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         DataField       =   "Address"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         DataField       =   "LastName"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         DataField       =   "Phone"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         DataField       =   "TechCode"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         DataField       =   "Active"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "First name:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last name:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tech code:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Active:"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
      Begin VB.ComboBox Combo1 
         DataField       =   "Station"
         DataSource      =   "Data3"
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "TechCode"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Station to repair:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pick a Technitian:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Technicians"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Shop.frx":E5C5
      Height          =   5415
      Left            =   6480
      TabIndex        =   8
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Edit"
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New order"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New order"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TechOrders"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSACAL.Calendar Calendar1 
      DataField       =   "ArriveDate"
      DataSource      =   "Data1"
      Height          =   2295
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   5055
      _Version        =   524288
      _ExtentX        =   8916
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   16777215
      Year            =   2011
      Month           =   6
      Day             =   13
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   4210752
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   8421504
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Shop.frx":E5D9
      Height          =   3615
      Left            =   720
      TabIndex        =   0
      Top             =   4800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   6120
      X2              =   6120
      Y1              =   720
      Y2              =   8400
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   630
      Left            =   -2588
      Picture         =   "Shop.frx":E5ED
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   2790
   End
End
Attribute VB_Name = "TechShed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Enabled = True
Command2.Visible = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
Data1.Recordset.AddNew
i = 1
End Sub

Private Sub Command10_Click()
If i = 1 Then
Data2.Recordset.Update
Else
Data2.Recordset.Update

Frame1.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command9.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
End Sub

Private Sub Command2_Click()
Frame1.Enabled = True
Command1.Visible = False
Command5.Enabled = False
Command6.Enabled = False
Command8.Enabled = False
Command10.Enabled = False
Data1.Recordset.Edit
i = 2
End Sub

Private Sub Command3_Click()
Frame1.Enabled = True
End Sub

Private Sub Command4_Click()
If i = 1 Then
Data1.Recordset.Cancel
Else
Data1.Recordset.CancelUpdate

Frame1.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command9.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
End Sub

Private Sub Command5_Click()
Frame2.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command9.Enabled = False
Command4.Enabled = False
Command6.Visible = False
Data2.Recordset.AddNew
i = 1
End Sub

Private Sub Command6_Click()
Frame2.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command9.Enabled = False
Command4.Enabled = False
Command5.Visible = False
Data2.Recordset.Edit
i = 2
End Sub
Private Sub Command8_Click()
If i = 1 Then
Data2.Recordset.Cancel
Else
Data2.Recordset.CancelUpdate

Frame1.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command9.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
End Sub

Private Sub Command9_Click()
If i = 1 Then
Data1.Recordset.Update
Else
Data1.Recordset.Update

Frame1.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command9.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
End Sub
