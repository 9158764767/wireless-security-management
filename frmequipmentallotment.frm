VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmequipmentallotment 
   Caption         =   "EQUIPMENT ALLOTMENT TO EMPLOYEE"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmequipmentallotment.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   15135
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmequipmentallotment.frx":66C82
      Height          =   1455
      Left            =   11760
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11880
      Top             =   4800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wireless security management.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wireless security management.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "EVENTS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtselectequipment 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   7200
      TabIndex        =   10
      Top             =   2520
      Width           =   3255
   End
   Begin VB.ComboBox cmbeallot 
      Height          =   315
      ItemData        =   "frmequipmentallotment.frx":66C97
      Left            =   7320
      List            =   "frmequipmentallotment.frx":66D31
      TabIndex        =   8
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EVENT REPORT"
      Height          =   975
      Left            =   6600
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txteququantity 
      DataField       =   "QUANTITY"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   7440
      TabIndex        =   4
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtempname 
      DataField       =   "ALLOTED EMPLOYEE"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   7920
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblselectequipment 
      Caption         =   "SELECTED EQUIPMENT"
      Height          =   735
      Left            =   2880
      TabIndex        =   9
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label lbleqyquantity 
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT QUANTITY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label lblequname 
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblempname 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "frmequipmentallotment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
frmrepairingsection.Show

End Sub

Private Sub cmdsave_Click()
txtselectequipment.Text = cmbeallot.Text
Adodc1.Recordset.Fields("ALLOTED EMPLOYEE") = txtempname.Text
Adodc1.Recordset.Fields("NAME") = txtselectequipment.Text
Adodc1.Recordset.Fields("QUANTITY") = txteququantity.Text







Adodc1.Recordset.Update

End Sub

Private Sub Command1_Click()
frm_allotmentbill.Show

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub
