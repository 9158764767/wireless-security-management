VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmtotalcostofsatellite 
   Caption         =   "TOTAL COST OF SATELLITE"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmtotalcostofsatellite.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   17325
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmtotalcostofsatellite.frx":66C82
      Height          =   2415
      Left            =   1440
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4260
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   3120
      Top             =   7800
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   $"frmtotalcostofsatellite.frx":66C97
      OLEDBString     =   $"frmtotalcostofsatellite.frx":66D2B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SATELLITE"
      Caption         =   "Adodc2"
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
   Begin VB.TextBox txtsatprice 
      DataField       =   "TOTAL COST"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   8760
      TabIndex        =   16
      Top             =   4680
      Width           =   4095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmtotalcostofsatellite.frx":66DBF
      Height          =   2655
      Left            =   1440
      TabIndex        =   15
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4683
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
      Left            =   3240
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
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
      Connect         =   $"frmtotalcostofsatellite.frx":66DD4
      OLEDBString     =   $"frmtotalcostofsatellite.frx":66E68
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
   Begin VB.CommandButton cmdbill 
      Caption         =   "BILL"
      Height          =   1215
      Left            =   10440
      TabIndex        =   14
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton cmdgolonch 
      Caption         =   "GO TO LONCH SECTION"
      Height          =   735
      Left            =   6120
      TabIndex        =   13
      Top             =   9840
      Width           =   7335
   End
   Begin VB.TextBox txtlaunchcost 
      Height          =   615
      Left            =   8880
      TabIndex        =   12
      Top             =   5520
      Width           =   4095
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
      Left            =   13080
      TabIndex        =   10
      Top             =   8040
      Width           =   1815
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
      Left            =   8280
      TabIndex        =   9
      Top             =   8160
      Width           =   1695
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
      Height          =   975
      Left            =   5280
      TabIndex        =   8
      Top             =   8160
      Width           =   2175
   End
   Begin VB.TextBox txtrocketcost 
      DataField       =   "PRICE"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   8760
      TabIndex        =   7
      Top             =   3360
      Width           =   4335
   End
   Begin VB.TextBox txtscost 
      Height          =   975
      Left            =   8880
      TabIndex        =   6
      Top             =   6240
      Width           =   4335
   End
   Begin VB.TextBox txtsprice 
      DataField       =   "PRICE"
      DataSource      =   "Adodc2"
      Height          =   855
      Left            =   8760
      TabIndex        =   5
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox txtsname 
      DataField       =   "NAME"
      DataSource      =   "Adodc2"
      Height          =   975
      Left            =   8760
      TabIndex        =   4
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label lblsatprice 
      Caption         =   "SATELIGHT PRICE"
      Height          =   615
      Left            =   4680
      TabIndex        =   17
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label lbllunch 
      Caption         =   "LAUNCHING COST"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblgain 
      BackStyle       =   0  'Transparent
      Caption         =   "ROCKET PRICE"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label lblscost 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL COST"
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
      TabIndex        =   2
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label lblsprice 
      BackStyle       =   0  'Transparent
      Caption         =   "SATELLITE PRICE"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblsname 
      BackStyle       =   0  'Transparent
      Caption         =   "SATELLITE NAME"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "frmtotalcostofsatellite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbill_Click()
frm_satbill.Show


End Sub

Private Sub cmdgolonch_Click()
frmlaunching.Show

End Sub

Private Sub cmdnext_Click()
frmnewfrequency.Show
End Sub

Private Sub cmdsave_Click()
Dim a, b, c, d As Double

a = txtsprice.Text
b = txtsatprice.Text
c = txtlaunchcost.Text
d = a + b + c
txtscost.Text = d



End Sub
