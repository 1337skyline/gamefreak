VERSION 5.00
Begin VB.Form GameStationsDetails 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   11685
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GameStationsDetails.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "GameFreakDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GameStations"
      Top             =   8400
      Visible         =   0   'False
      Width           =   2700
   End
End
Attribute VB_Name = "GameStationsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
