VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmnewoffice 
   BackColor       =   &H00FF8080&
   Caption         =   "NEW OFFICE REGISTRATION"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmnewoffice.frx":0000
   ScaleHeight     =   10080
   ScaleWidth      =   15525
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmnewoffice.frx":66C82
      Height          =   1935
      Left            =   360
      TabIndex        =   15
      Top             =   4440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3413
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
      Left            =   840
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "OFFICE"
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
   Begin VB.TextBox txtarea 
      DataField       =   "UNDERTAKING AREA"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   8280
      TabIndex        =   14
      Top             =   7800
      Width           =   3975
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
      Height          =   1095
      Left            =   10080
      TabIndex        =   12
      Top             =   9120
      Width           =   1935
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
      Height          =   1095
      Left            =   7680
      TabIndex        =   11
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton cmdregister 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      TabIndex        =   10
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtcontact 
      DataField       =   "CONTACT"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8160
      TabIndex        =   9
      Top             =   6600
      Width           =   3375
   End
   Begin VB.TextBox txtincharge 
      DataField       =   "INCHARGE"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8160
      TabIndex        =   7
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox txtlocation 
      DataField       =   "LOCATION"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8160
      TabIndex        =   5
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox txtoname 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8160
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8160
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "undertaking areas"
      Height          =   855
      Left            =   4680
      TabIndex        =   13
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblcontact 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblincharge 
      BackStyle       =   0  'Transparent
      Caption         =   "INCHARGE"
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
      Left            =   5400
      TabIndex        =   6
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label lbllocation 
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lbloname 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
      Left            =   5400
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
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
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmnewoffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
frmnewemployee.Show

End Sub

Private Sub cmdregister_Click()
Adodc1.Recordset.Fields("ID") = txtid.Text
Adodc1.Recordset.Fields("NAME") = txtoname.Text
Adodc1.Recordset.Fields("LOCATION") = txtlocation.Text
Adodc1.Recordset.Fields("INCHARGE") = txtincharge.Text
Adodc1.Recordset.Fields("CONTACT") = txtcontact.Text
Adodc1.Recordset.Fields("UNDERTAKING AREA") = txtarea.Text

Adodc1.Recordset.Update
MsgBox ("data added successfully")


End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub
