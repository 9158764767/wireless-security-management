VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmfuelrequirement 
   Caption         =   "FUEL REQUIREMENT"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmfuelrequirement.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   16425
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmfuelrequirement.frx":66C82
      Height          =   1455
      Left            =   1920
      TabIndex        =   11
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   855
      Left            =   2520
      Top             =   3120
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1508
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
      Connect         =   $"frmfuelrequirement.frx":66C97
      OLEDBString     =   $"frmfuelrequirement.frx":66D2B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ROCKET"
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
   Begin VB.TextBox txtfuelrequire 
      DataField       =   "FUEL AMOUNT"
      DataSource      =   "Adodc1"
      Height          =   765
      Left            =   8640
      TabIndex        =   10
      Top             =   2040
      Width           =   3375
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
      Height          =   1095
      Left            =   10200
      TabIndex        =   9
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
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
      TabIndex        =   8
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "BILL"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtfuelprice 
      DataField       =   "PRICE"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8640
      TabIndex        =   6
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox txtfuelname 
      DataField       =   "FUELNAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8640
      TabIndex        =   4
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtrocketname 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8640
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblfuelprice 
      BackStyle       =   0  'Transparent
      Caption         =   "FUEL PRICE"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label lblfuelname 
      BackStyle       =   0  'Transparent
      Caption         =   "FUEL NAME"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblfuelrequire 
      BackStyle       =   0  'Transparent
      Caption         =   "FUEL REQUIRED"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblrocketname 
      BackStyle       =   0  'Transparent
      Caption         =   "ROCKET NAME"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "frmfuelrequirement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
frmrocketselection.Show

End Sub

Private Sub cmdsave_Click()
frm_fuelbill.Show


End Sub
