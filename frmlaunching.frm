VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmlaunching 
   Caption         =   "LAUNCHING OF SATELLITE"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmlaunching.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   15600
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmlaunching.frx":66C82
      Height          =   1575
      Left            =   1080
      TabIndex        =   16
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2778
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
      Left            =   1080
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Connect         =   $"frmlaunching.frx":66C97
      OLEDBString     =   $"frmlaunching.frx":66D2B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "LOUNCH"
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "LAUNCHING SECTION"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      TabIndex        =   11
      Top             =   360
      Width           =   8895
      Begin VB.OptionButton opts1 
         BackColor       =   &H8000000D&
         Caption         =   "STATION 1"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton opts2 
         BackColor       =   &H8000000D&
         Caption         =   "STATION 2"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
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
      Height          =   855
      Left            =   7200
      TabIndex        =   10
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "SAVE DETAILS"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton lbllaunch 
      Caption         =   "LAUNCH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6840
      TabIndex        =   8
      Top             =   7080
      Width           =   3375
   End
   Begin VB.TextBox txtoperationcontrol 
      DataField       =   "OPTHEAD"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8640
      TabIndex        =   7
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13320
      Top             =   480
   End
   Begin VB.TextBox txtlocation 
      DataField       =   "LOCATION"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8640
      TabIndex        =   5
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox txtdate 
      DataField       =   "DATE"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8640
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox txtsettime 
      DataField       =   "STATION"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   6360
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   12960
      TabIndex        =   15
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbloperationcontrol 
      BackStyle       =   0  'Transparent
      Caption         =   "OPERATION CONTROLLER"
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
      TabIndex        =   6
      Top             =   5520
      Width           =   3975
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
      Left            =   4320
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lbldate 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lblsettime 
      BackStyle       =   0  'Transparent
      Caption         =   "SET TIME"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "frmlaunching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
Adodc1.Recordset.Fields("DATE") = txtdate.Text
Adodc1.Recordset.Fields("LOCATION") = txtlocation.Text
Adodc1.Recordset.Fields("OPTHEAD") = txtoperationcontrol.Text
MsgBox ("UPDATED")
If txtdate.Text = "" And txtlocation.Text = "" And txtoperationcontrol.Text = "" Then
MsgBox ("PLEASE FILLUP DETAILS FOR LOUNCHING*********")
End If
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
DataGrid1.Visible = False
End Sub

Private Sub lbllaunch_Click()
Timer1.Enabled = True
ProgressBar1.Visible = True
Label1.Visible = True
End Sub

Private Sub opts1_Click()
txtsettime.Text = "HUSTAN-USA"

End Sub

Private Sub opts2_Click()
txtsettime.Text = "SHREE HARI KOTTA-INDIA"
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label1.Caption = ProgressBar1.Value & "%TOGO"
If ProgressBar1.Value = 100 Then
Timer1.Enabled = False
Timer1.Enabled = False
ProgressBar1.Visible = False
Label1.Visible = False
MsgBox ("THE LOUNCHING PROCESS IS REGISTER ON: " + txtdate.Text + "AT:" + txtlocation.Text + "OFFICER:" + txtoperationcontrol.Text)
End If

End Sub
