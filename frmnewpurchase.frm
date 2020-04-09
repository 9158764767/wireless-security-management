VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmnewpurchase 
   BackColor       =   &H00FF8080&
   Caption         =   "NEW PURCHASE OF EQUIPMENT"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmnewpurchase.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   17160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "BILL"
      Height          =   1215
      Left            =   11280
      TabIndex        =   15
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtcompany 
      DataField       =   "EQUIPMENT COMPANY"
      DataSource      =   "Adodc1"
      Height          =   1095
      Left            =   3720
      TabIndex        =   14
      Top             =   6000
      Width           =   3135
   End
   Begin VB.ComboBox cmbclist 
      Height          =   315
      ItemData        =   "frmnewpurchase.frx":66C82
      Left            =   4320
      List            =   "frmnewpurchase.frx":66C9E
      TabIndex        =   12
      Top             =   4920
      Width           =   3135
   End
   Begin VB.ComboBox cmbelist 
      Height          =   315
      ItemData        =   "frmnewpurchase.frx":66CE3
      Left            =   4320
      List            =   "frmnewpurchase.frx":66D7D
      TabIndex        =   11
      Top             =   2280
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmnewpurchase.frx":67006
      Height          =   1455
      Left            =   12000
      TabIndex        =   10
      Top             =   5400
      Width           =   3735
      _ExtentX        =   6588
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
      Height          =   975
      Left            =   8880
      Top             =   5640
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
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
      RecordSource    =   "NewEquip"
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
   Begin VB.TextBox txtename 
      DataField       =   "EQUIPMENT NAME"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   4200
      TabIndex        =   9
      Top             =   3360
      Width           =   3735
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
      Height          =   1215
      Left            =   14160
      TabIndex        =   6
      Top             =   3240
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
      Height          =   1215
      Left            =   8880
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txteqprice 
      DataField       =   "EQUIPMENT COST"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   11880
      TabIndex        =   4
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblcompany 
      Caption         =   "COMPANY SELECTED"
      Height          =   855
      Left            =   600
      TabIndex        =   13
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label lblid 
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
      Height          =   855
      Left            =   1320
      TabIndex        =   8
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblcname 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY LIST"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label lbleqprice 
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT PRICE"
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
      Left            =   8640
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lbleqlist 
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT LIST"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lbleqname 
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
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   3480
      Width           =   3015
   End
End
Attribute VB_Name = "frmnewpurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
frmeventadnoperation.Show

End Sub

Private Sub cmdsave_Click()
txtename.Text = cmbelist.Text
txtcompany.Text = cmbclist.Text
Adodc1.Recordset.Fields("ID") = txtid.Text
Adodc1.Recordset.Fields("EQUIPMENT NAME") = txtename.Text
Adodc1.Recordset.Fields("EQUIPMENT COMPANY") = txtcompany.Text
Adodc1.Recordset.Fields("EQUIPMENT COST") = txteqprice.Text





Adodc1.Recordset.Update



End Sub

Private Sub Command1_Click()
frm_newpurchasebill.Show

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub
