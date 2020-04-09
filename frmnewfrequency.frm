VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmnewfrequency 
   Caption         =   "new frequency"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmnewfrequency.frx":0000
   ScaleHeight     =   9615
   ScaleWidth      =   16425
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "BILL"
      Height          =   1095
      Left            =   12240
      TabIndex        =   12
      Top             =   7680
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmnewfrequency.frx":66C82
      Height          =   1455
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
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
      Height          =   330
      Left            =   3840
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
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
      Connect         =   $"frmnewfrequency.frx":66C97
      OLEDBString     =   $"frmnewfrequency.frx":66D2B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "FREQUENCY"
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
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10200
      TabIndex        =   10
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   9
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdproceed 
      Caption         =   "PROCED"
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
      Left            =   5280
      TabIndex        =   8
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox txtcost 
      DataField       =   "FCOST"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   9960
      TabIndex        =   7
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox txtsetf 
      DataField       =   "FREQUENCYNAME"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   9960
      TabIndex        =   5
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtchannelset 
      DataField       =   "CHANNAL"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   9960
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox txtid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   9960
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblsetf 
      BackStyle       =   0  'Transparent
      Caption         =   "FREQUENCY NAME"
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
      Left            =   6120
      TabIndex        =   6
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label lblcost 
      BackStyle       =   0  'Transparent
      Caption         =   "COST"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      TabIndex        =   4
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label lblchannelset 
      BackStyle       =   0  'Transparent
      Caption         =   "CHANNEL SET"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
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
      Left            =   6240
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "frmnewfrequency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
txtid.Text = ""
txtchannelset.Text = ""
txtsetf.Text = ""
txtcost.Text = ""

End Sub

Private Sub cmdexit_Click()
mdihomepage.Show
End Sub

Private Sub cmdproceed_Click()
If txtid.Text = "" And txtchannelset.Text = "" Then
MsgBox ("PLEASE ENTER THE FIELDS*****")
ElseIf txtsetf.Text = "" And txtcost.Text = "" Then
MsgBox ("PLEASE ENTER THE FIELDS*****")
End If


'Adodc1.Recordset.Fields("ID") = txtid.Text
Adodc1.Recordset.Fields("CHANNAL") = txtchannelset.Text
Adodc1.Recordset.Fields("FREQUENCYNAME") = txtsetf.Text
Adodc1.Recordset.Fields("FCOST") = txtcost.Text
End Sub

Private Sub Command1_Click()
frm_FREQBILL.Show


End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew



End Sub
