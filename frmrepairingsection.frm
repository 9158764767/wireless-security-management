VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmrepairingsection 
   BackColor       =   &H00FF8080&
   Caption         =   "REPAIRING SECTION"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmrepairingsection.frx":0000
   ScaleHeight     =   9600
   ScaleWidth      =   17055
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmrepairingsection.frx":66C82
      Height          =   1695
      Left            =   720
      TabIndex        =   17
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
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
      Height          =   615
      Left            =   360
      Top             =   6000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
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
      RecordSource    =   "REPAIR"
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
   Begin VB.TextBox txtoprice 
      DataField       =   "ORIGINAL PRICE"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8760
      TabIndex        =   15
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txtname 
      DataField       =   "date of dishcharge"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8280
      TabIndex        =   14
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   12
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   11
      Top             =   8640
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   10
      Top             =   8640
      Width           =   1695
   End
   Begin VB.TextBox txtdod 
      DataField       =   "date of dishcharge"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8520
      TabIndex        =   9
      Top             =   7320
      Width           =   3615
   End
   Begin VB.TextBox txtdoa 
      DataField       =   "date of admit"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   8280
      TabIndex        =   7
      Top             =   6000
      Width           =   3615
   End
   Begin VB.TextBox txttotalprice 
      DataField       =   "COST"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   8760
      TabIndex        =   5
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox txtfault 
      DataField       =   "FAULT"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   8280
      TabIndex        =   3
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox txtid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   8280
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "ORIGINAL PRICE"
      Height          =   735
      Left            =   4200
      TabIndex        =   16
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Label lblid 
      BackStyle       =   0  'Transparent
      Caption         =   "name"
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
      Left            =   4320
      TabIndex        =   13
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lbldod 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF DISHCHARGE"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Label lbldoa 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF ADMITTED"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label lbltotalprice 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL REPAIRING PRICE"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label lblfault 
      BackStyle       =   0  'Transparent
      Caption         =   "fault"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label lbleqname 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "frmrepairingsection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
frmtotallossandgainbyequipment.Show

End Sub

Private Sub cmdsave_Click()
Adodc1.Recordset.Fields("ID") = txtid.Text
Adodc1.Recordset.Fields("EQUIPMENT NAME") = txtname.Text
Adodc1.Recordset.Fields("FAULT REPAIRED") = txtfault.Text
Adodc1.Recordset.Fields("COST") = txttotalprice.Text
Adodc1.Recordset.Fields("ORIGINAL PRICE") = txtoprice.Text
Adodc1.Recordset.Fields("date of admit") = txtdoa.Text
Adodc1.Recordset.Fields("date of dishcharge") = txtdod.Text

Adodc1.Recordset.Update

MsgBox ("Data Added Succefully")


End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub
