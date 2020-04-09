VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmtotallossandgainbyequipment 
   Caption         =   "TOTAL LOSS AND GAIN BY EQUIPMENT"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmtotallossandgainbyequipment.frx":0000
   ScaleHeight     =   9630
   ScaleWidth      =   16830
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "back"
      Height          =   615
      Left            =   11040
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "CALCULATE"
      Height          =   855
      Left            =   10800
      TabIndex        =   17
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox txtrepairprice 
      DataField       =   "COST"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   2880
      TabIndex        =   16
      Top             =   4560
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "REPAIRED"
      Height          =   615
      Left            =   10800
      TabIndex        =   14
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DESTROYED"
      Height          =   615
      Left            =   8040
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BILL"
      Height          =   1095
      Left            =   7080
      TabIndex        =   12
      Top             =   5280
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmtotallossandgainbyequipment.frx":66C82
      Height          =   1095
      Left            =   480
      TabIndex        =   11
      Top             =   6480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1931
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
      Left            =   4920
      Top             =   7080
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.TextBox txtfaultresult 
      DataField       =   "FAULT"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10920
      TabIndex        =   10
      Top             =   1920
      Width           =   2655
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
      TabIndex        =   9
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtloss 
      DataField       =   "TOTAL LOSS"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   10320
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtprice 
      DataField       =   "ORIGINAL PRICE"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   3360
      TabIndex        =   5
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox txtquantity 
      DataField       =   "REPAIR QUANTITY"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txteqname 
      DataField       =   "EQUIPMENT NAME"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "REPAIRING PRICE"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lblgain 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL LOSS"
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
      Left            =   6840
      TabIndex        =   7
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblcostofeq 
      BackStyle       =   0  'Transparent
      Caption         =   "FAULT REPAIRED"
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
      Left            =   7200
      TabIndex        =   6
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblprice 
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
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblquantity 
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT QUANTITY"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
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
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmtotallossandgainbyequipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmrepairingsection.Show

End Sub

Private Sub cmdcalculate_Click()
txtloss.Text = Val(txtrepairprice.Text) + Val(txtprice.Text)

MsgBox ("repair is failed")

End Sub

Private Sub cmdnext_Click()

End Sub

Private Sub cmdsave_Click()
Adodc1.Recordset.Update

Adodc1.Recordset.Fields("TOTAL LOSS") = txtloss.Text
Adodc1.Recordset.Fields("FAULT REPAIRED") = txtfaultresult.Text
MsgBox ("Record Added Succefully")


End Sub

Private Sub Command1_Click()
frm_repairbill.Show

End Sub


Private Sub Command2_Click()
frmrepairingsection.Show

End Sub

Private Sub Option1_Click()
txtloss.Visible = True
txtfaultresult.Text = Option1.Caption

End Sub

Private Sub Option2_Click()
txtfaultresult.Text = Option2.Caption
txtloss.Visible = False

End Sub
