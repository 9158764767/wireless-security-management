VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmrocketselection 
   Caption         =   "ROCKET SELECTION"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmrocketselection.frx":0000
   ScaleHeight     =   9960
   ScaleWidth      =   16275
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtcompprice 
      Height          =   495
      Left            =   14520
      TabIndex        =   21
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtpartptice 
      Height          =   495
      Left            =   14520
      TabIndex        =   20
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   615
      Left            =   13200
      TabIndex        =   19
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox txtcomp 
      Height          =   495
      Left            =   10800
      TabIndex        =   16
      Top             =   5760
      Width           =   3015
   End
   Begin VB.TextBox txtpart 
      Height          =   375
      Left            =   11280
      TabIndex        =   15
      Top             =   4560
      Width           =   2655
   End
   Begin VB.ComboBox cmbcomp 
      Height          =   315
      ItemData        =   "frmrocketselection.frx":66C82
      Left            =   10800
      List            =   "frmrocketselection.frx":66C9E
      TabIndex        =   14
      Top             =   5160
      Width           =   3015
   End
   Begin VB.ComboBox cmbparts 
      Height          =   315
      ItemData        =   "frmrocketselection.frx":66D26
      Left            =   11040
      List            =   "frmrocketselection.frx":66D45
      TabIndex        =   13
      Top             =   3840
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmrocketselection.frx":66DAD
      Height          =   1935
      Left            =   360
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   495
      Left            =   840
      Top             =   6240
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   $"frmrocketselection.frx":66DC2
      OLEDBString     =   $"frmrocketselection.frx":66E56
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
      Left            =   9240
      TabIndex        =   10
      Top             =   7680
      Width           =   1695
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
      Left            =   5640
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txtfuelamt 
      DataField       =   "FUEL AMOUNT"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txtcost 
      DataField       =   "PRICE"
      DataSource      =   "Adodc1"
      Height          =   1005
      Left            =   11040
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtrname 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   5040
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
   Begin VB.ComboBox cmbtype 
      DataField       =   "TYPE"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frmrocketselection.frx":66EEA
      Left            =   8040
      List            =   "frmrocketselection.frx":66F00
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   11520
      TabIndex        =   22
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "COMPANY SELECTED"
      Height          =   615
      Left            =   8520
      TabIndex        =   18
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "PART SELECTED"
      Height          =   375
      Left            =   8400
      TabIndex        =   17
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "COMPANY LIST"
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblfuelamt 
      BackStyle       =   0  'Transparent
      Caption         =   "FUEL AMOUNT AS SELECTED"
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
      Left            =   480
      TabIndex        =   8
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Label lblother 
      BackStyle       =   0  'Transparent
      Caption         =   "PARTS LIST"
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
      Left            =   8520
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label lbladdress 
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE"
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
      Left            =   8520
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblrname 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label lbltype 
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
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
      Left            =   5400
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmrocketselection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdnext_Click()
If txtrname.Text = " " And txtfuelamt.Text = "" Then
MsgBox ("PLEASE ENTER THE RECORDS TO PROCESS")
End If
If txtcost.Text = "" And cmbtype.Text = "" Then
MsgBox ("PLEASE ENTER THE RECORDS TO PROCESS")
End If
Adodc1.Recordset.Fields("NAME") = txtrname.Text
Adodc1.Recordset.Fields("TYPE") = cmbtype.Text
Adodc1.Recordset.Fields("FUEL AMOUNT") = txtfuelamt.Text
Adodc1.Recordset.Fields("PARTS") = txtpart.Text
Adodc1.Recordset.Fields("COMPANY") = txtcomp.Text
Adodc1.Recordset.Fields("PRICE") = txtcost.Text
MsgBox ("SUCESS FULL*****")
frmtotalcostofsatellite.Show
Adodc1.Recordset.Update


End Sub

Private Sub Command1_Click()

  txtpart.Text = cmbparts.Text
 If txtpart.Text = "NOSE CONES" Then
 txtpartptice.Text = 200
 End If
 If txtpart.Text = "PAYLOAD System" Then
 txtpartptice.Text = 2000
 End If
 If txtpart.Text = "GUIDANCE System" Then
 txtpartptice.Text = 3000
 End If
 If txtpart.Text = "FUEL INJECTORS" Then
 txtpartptice.Text = 400000
 End If
 If txtpart.Text = "FRAMES" Then
 txtpartptice.Text = 5000
 End If
 If txtpart.Text = "OXIDISER" Then
 txtpartptice.Text = 6000000
 End If
 If txtpart.Text = "PUMPS" Then
 txtpartptice.Text = 7000
 End If
 If txtpart.Text = "NOZZEL" Then
 txtpartptice.Text = 8000000
 End If
 If txtpart.Text = "FINS" Then
 txtpartptice.Text = 900
 End If
 If txtpart.Text = "Rocket Fuel Engictor" Then
 txtpartptice.Text = 1000
 End If

 
  txtcomp.Text = cmbcomp.Text
 
 If txtcomp.Text = "BOING" Then
 txtcompprice.Text = 10000
 End If
 If txtcomp.Text = "ORBITAL SCIENCE" Then
 txtcompprice.Text = 20000
 End If
 If txtcomp.Text = "STERRA NEVADA CORPORATION" Then
  txtcompprice.Text = 300
 End If
 If txtcomp.Text = "VIRGIN GALACTIC" Then
 txtcompprice.Text = 4000000
 End If
 If txtcomp.Text = "XCOR AEROSPEACE" Then
  txtcompprice.Text = 500
 End If
 If txtcomp.Text = "ADASTRA" Then
   txtcompprice.Text = 6000
 End If
 If txtcomp.Text = "" Then
 txtcompprice.Text = 7000
 End If
 If txtcomp.Text = "NSSC GLOBAL.LTD" Then
 txtcompprice.Text = 800000
 End If
 If txtcomp.Text = "VIA SAT" Then
 txtcompprice.Text = 1000
 End If
 
 

txtcost.Text = Val(txtcompprice.Text) + Val(txteqprice.Text)
 
End Sub


 
Private Sub Form_Load()
Adodc1.Recordset.AddNew



End Sub
