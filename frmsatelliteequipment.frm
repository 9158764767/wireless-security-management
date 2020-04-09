VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsatelliteequipment 
   Caption         =   "SATELLITE EQUIPMENT SELECTION"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmsatelliteequipment.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   18735
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   855
      Left            =   11280
      TabIndex        =   23
      Top             =   4920
      Width           =   975
   End
   Begin VB.ComboBox cmbcomp 
      Height          =   315
      ItemData        =   "frmsatelliteequipment.frx":66C82
      Left            =   8160
      List            =   "frmsatelliteequipment.frx":66CA4
      TabIndex        =   22
      Text            =   "-SELECT COMP"
      Top             =   5400
      Width           =   2535
   End
   Begin VB.ComboBox cmbeqip 
      Height          =   315
      ItemData        =   "frmsatelliteequipment.frx":66D60
      Left            =   8160
      List            =   "frmsatelliteequipment.frx":66D7F
      TabIndex        =   21
      Text            =   "-SELEECT EQ"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtcomprice 
      Height          =   615
      Left            =   9360
      TabIndex        =   20
      Top             =   6600
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   1200
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "SATELLITE"
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
   Begin VB.TextBox txtcomp 
      DataField       =   "COMPANY"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   13920
      TabIndex        =   18
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txteq 
      DataField       =   "EQIPMENT"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   14040
      TabIndex        =   17
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "TOTAL COST"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   2640
      TabIndex        =   15
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton cmdprice 
      Caption         =   "CALCULET PRICE"
      Height          =   735
      Left            =   10680
      TabIndex        =   14
      Top             =   6720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmsatelliteequipment.frx":66E5E
      Height          =   1575
      Left            =   720
      TabIndex        =   13
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   11400
      TabIndex        =   12
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdbill 
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
      Height          =   855
      Left            =   9240
      TabIndex        =   11
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H8000000D&
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
      Left            =   7080
      TabIndex        =   10
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox txteqprice 
      DataField       =   "EQIPMENT"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   6840
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtvfuel 
      Height          =   855
      Left            =   8160
      TabIndex        =   8
      Top             =   3240
      Width           =   3975
   End
   Begin VB.TextBox txtorbit 
      DataField       =   "ORBIT"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8160
      TabIndex        =   7
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtstype 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8160
      TabIndex        =   6
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "COMPANY"
      Height          =   255
      Left            =   12960
      TabIndex        =   25
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "EQUIPMENT"
      Height          =   255
      Left            =   12840
      TabIndex        =   24
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "COMPRICE"
      Height          =   615
      Left            =   8280
      TabIndex        =   19
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label lbltotal 
      Caption         =   "TOTALCOST"
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblprice 
      BackStyle       =   0  'Transparent
      Caption         =   "EQPRICE"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label lblcompany 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lbletype 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label lblvfuel 
      BackStyle       =   0  'Transparent
      Caption         =   "V FUEL"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblorbit 
      BackStyle       =   0  'Transparent
      Caption         =   "ORBIT"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblstype 
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
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "frmsatelliteequipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbill_Click()
frm_satbill.Show


End Sub

Private Sub cmdnext_Click()
frmrocketselection.Show

End Sub

Private Sub cmdprice_Click()

Text1.Text = Val(txtcomprice) + Val(txteqprice)




End Sub

Private Sub Command1_Click()

  txteq.Text = cmbeqip.Text
 If txteq.Text = "Tracking Telemetry Command & Ranging(TTC&R)" Then
 txteqprice.Text = 2000000
 End If
 If txteq.Text = "BATTERIES" Then
 txteqprice.Text = 3000000
 End If
 If txteq.Text = "Reaction Control System" Then
 txteqprice.Text = 400000
 End If
 If txteq.Text = "Speacecraft Control Processor" Then
 txteqprice.Text = 500000
 End If
 If txteq.Text = "Thermal Controlar" Then
 txteqprice.Text = 6000000
 End If
 If txteq.Text = "Command Antenna" Then
 txteqprice.Text = 7000000
 End If
 If txteq.Text = "Communication Antenna" Then
 txteqprice.Text = 8000000
 End If
 If txteq.Text = "R.F.Reciver &Transmittor" Then
 txteqprice.Text = 9000000
 End If
 If txteq.Text = "Rocket Fuel Engictor" Then
 txteqprice.Text = 10000000
 End If
 
 txtcomp.Text = cmbcomp.Text
 
 If txtcomp.Text = "Echo Star Communication.Corporation" Then
 txtcomprice.Text = 10000
 End If
 If txtcomp.Text = "XM Satellite Systems.INC" Then
 txtcomprice.Text = 20000
 End If
 If txtcomp.Text = "Boing Satellite Systems.INCPHILCO FORD" Then
 txtcomprice.Text = 3000000
 End If
 If txtcomp.Text = "QINETIQ SPEACE N-V" Then
 txtcomprice.Text = 4000000
 End If
 If txtcomp.Text = "RCA ASTRO" Then
 txtcomprice.Text = 5000000
 End If
 If txtcomp.Text = "ROCK WHELL" Then
 txtcomprice.Text = 6000000
 End If
 If txtcomp.Text = "ISAT.LTD" Then
 txtcomprice.Text = 7000000
 End If
 If txtcomp.Text = "NSSC GLOBAL.LTD" Then
 txtcomprice.Text = 800000
 End If
 If txtcomp.Text = "VIA SAT" Then
 txtcomprice.Text = 1000000
 End If
 
 
 
 
 
 
 
 
 
 
 
 
 
 
End Sub

