VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmequipmentusesinoperation 
   BackColor       =   &H00FF8080&
   Caption         =   "EQUIPMENT USES IN OPERATION"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmequipmentusesinoperation.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   15465
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "CALCULATE"
      Height          =   855
      Left            =   11880
      TabIndex        =   19
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox txtno 
      DataField       =   "QUANTITY"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   9720
      TabIndex        =   18
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   735
      Left            =   5040
      TabIndex        =   17
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox txtoselected 
      DataField       =   "OPERATION NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9600
      TabIndex        =   16
      Top             =   1560
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmequipmentusesinoperation.frx":66C82
      Height          =   1575
      Left            =   1320
      TabIndex        =   15
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
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
      Height          =   330
      Left            =   1080
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.TextBox txttotalcost 
      DataField       =   "TOTAL COST"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8880
      TabIndex        =   14
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox txteprice 
      DataField       =   "EQUIPMENT PRICE"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   8760
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtoprice 
      DataField       =   "OPERATION PRICE"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BILL"
      Height          =   975
      Left            =   2760
      TabIndex        =   8
      Top             =   7200
      Width           =   1815
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
      Left            =   9840
      TabIndex        =   7
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdnext2 
      BackColor       =   &H00FFFFFF&
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
      Height          =   975
      Left            =   7080
      TabIndex        =   6
      Top             =   7320
      Width           =   1815
   End
   Begin VB.TextBox txtequipmentselected 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9600
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.ComboBox cmboperationlist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmequipmentusesinoperation.frx":66C97
      Left            =   9720
      List            =   "frmequipmentusesinoperation.frx":66CAD
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "TOTAL COST"
      Height          =   855
      Left            =   5280
      TabIndex        =   13
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "EQUIPMENT PRICE"
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lbloperationprice 
      Caption         =   "OPERATION PRICE"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lblnoofequipment 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF EQUIPMENT"
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
      TabIndex        =   5
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblequipmentselected 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "EQUIPMENT SELECTED"
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
      TabIndex        =   3
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lbloperation 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "OPERATION SELECTED"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lbloperationlist 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "OPERATION LIST"
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
      Left            =   5520
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frmequipmentusesinoperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext2_Click()
frmeventadnoperation.Show

End Sub

Private Sub cmdsave_Click()
Adodc1.Recordset.Fields("OPERATION NAME") = txtoselected.Text
Adodc1.Recordset.Fields("QUANTITY") = txtno.Text
Adodc1.Recordset.Fields("OPERATION PRICE") = txtoprice.Text
Adodc1.Recordset.Fields("EQUIPMENT PRICE") = txteprice.Text
Adodc1.Recordset.Fields("TOTAL COST") = txttotalcost.Text



Adodc1.Recordset.Update

 txtoselected.Text = cmboperationlist.Text
 
 

End Sub

Private Sub Command1_Click()
frm_equipmentbill.Show

End Sub

Private Sub Command2_Click()
txttotalcost.Text = Val(txtoprice.Text) + Val(txteprice.Text)
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub
