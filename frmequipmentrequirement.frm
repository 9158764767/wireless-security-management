VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmequipmentrequirement 
   BackColor       =   &H00FF8080&
   Caption         =   "EQUIPMENT REQUIREMENT"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16095
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000016&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmequipmentrequirement.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   16095
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmequipmentrequirement.frx":66C82
      Height          =   1695
      Left            =   11400
      TabIndex        =   12
      Top             =   2280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Height          =   495
      Left            =   11760
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
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
      RecordSource    =   "ground"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtcompany 
      DataField       =   "COMPANY"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   6240
      TabIndex        =   11
      Top             =   3840
      Width           =   2415
   End
   Begin VB.ComboBox cmbcompany 
      Height          =   360
      ItemData        =   "frmequipmentrequirement.frx":66C97
      Left            =   6240
      List            =   "frmequipmentrequirement.frx":66CB3
      TabIndex        =   10
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txteselect 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   6240
      TabIndex        =   9
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   1095
      Left            =   4200
      TabIndex        =   6
      Top             =   5520
      Width           =   1935
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
      Height          =   975
      Left            =   8640
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdnext1 
      BackColor       =   &H00C0C0FF&
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
      Height          =   975
      Left            =   6480
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox cmbequipmentlist 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmequipmentrequirement.frx":66CF8
      Left            =   6120
      List            =   "frmequipmentrequirement.frx":66D92
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblcompany 
      Caption         =   "company selected"
      Height          =   735
      Left            =   2760
      TabIndex        =   8
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "company list"
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblequipmentselected 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT SELECTED"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lblequipmentlist 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT LIST"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblequipmentrequirement 
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT REQUIREMENT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   6360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmequipmentrequirement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmequipmentrequirement.Hide

End Sub

Private Sub cmdnext1_Click()
frmequipmentusesinoperation.Show

End Sub

Private Sub Command1_Click()
txteselect.Text = cmbequipmentlist.Text
txtcompany.Text = cmbcompany.Text
Adodc1.Recordset.Update

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub
