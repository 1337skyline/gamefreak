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
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "F:\Greys.Anatomy.S07\VisualBasic 6.0\game-freak\trunk\Access\GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Suppliers"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10440
      TabIndex        =   25
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      TabIndex        =   23
      Text            =   "Text5"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10440
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10440
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1920
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Shop.frx":E5C5
      Height          =   5415
      Left            =   6480
      TabIndex        =   14
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10560
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Update"
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New order"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New order"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1815
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
      RecordSource    =   "Orders"
      Top             =   8520
      Width           =   1815
   End
   Begin MSACAL.Calendar Calendar1 
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
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Active:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9600
      TabIndex        =   26
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tech code:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   24
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9600
      TabIndex        =   22
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last name:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9600
      TabIndex        =   21
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First name:"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      Top             =   1920
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   6120
      X2              =   6120
      Y1              =   720
      Y2              =   8400
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pick a Technitian:"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Station to repair:"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
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
