VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmeventadnoperation 
   BackColor       =   &H00FF8080&
   Caption         =   "EVENTS AND OPERATION"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15705
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmeventadnoperation.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   15705
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "CALCULATE"
      Height          =   735
      Left            =   12840
      TabIndex        =   21
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8520
      TabIndex        =   20
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtlocation 
      DataField       =   "LOCATION"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8880
      TabIndex        =   18
      Top             =   3480
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmeventadnoperation.frx":66C82
      Height          =   1935
      Left            =   1800
      TabIndex        =   16
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   855
      Left            =   1920
      Top             =   7200
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wireless security management.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\wireless security management.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "EVENTS"
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
   Begin VB.ComboBox cmbevent 
      Height          =   315
      ItemData        =   "frmeventadnoperation.frx":66C97
      Left            =   8760
      List            =   "frmeventadnoperation.frx":66CB0
      TabIndex        =   15
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton cmdbill 
      Caption         =   "BILL"
      Height          =   1095
      Left            =   12840
      TabIndex        =   14
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox txttotalcost 
      DataField       =   "TOTAL COST"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   8760
      TabIndex        =   13
      Top             =   8400
      Width           =   2535
   End
   Begin VB.TextBox txtcostofoperation 
      DataField       =   "COST OF OPERATION"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8760
      TabIndex        =   11
      Top             =   7440
      Width           =   2895
   End
   Begin VB.TextBox txtoselected 
      DataField       =   "OPERATION SELECTED"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   8880
      TabIndex        =   9
      Top             =   4560
      Width           =   2895
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
      Height          =   855
      Left            =   12960
      TabIndex        =   7
      Top             =   4200
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
      Height          =   855
      Left            =   13080
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
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
      Height          =   855
      Left            =   13080
      TabIndex        =   5
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txtcostofevent 
      DataField       =   "COST OF EVENT"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8760
      TabIndex        =   4
      Top             =   6360
      Width           =   3375
   End
   Begin VB.TextBox txteventselected 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   8640
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label lblid 
      Caption         =   "ID"
      Height          =   735
      Left            =   5760
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "LOCATION"
      Height          =   855
      Left            =   5280
      TabIndex        =   17
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "TOTAL COST"
      Height          =   855
      Left            =   5520
      TabIndex        =   12
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "COST OF OPERATION"
      Height          =   855
      Left            =   5040
      TabIndex        =   10
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "COST OF EVENT"
      Height          =   975
      Left            =   5280
      TabIndex        =   8
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label lbloperationselected 
      BackStyle       =   0  'Transparent
      Caption         =   "OPERATION NAME"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label lbleventselected 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT EVENT"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lbleventname 
      BackStyle       =   0  'Transparent
      Caption         =   "EVENT NAME"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmeventadnoperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbill_Click()
frm_eventbill.Show

End Sub

Private Sub cmdcalculate_Click()
txttotalcost.Text = Val(txtcostofevent.Text) + Val(txtcostofoperation.Text)

End Sub

Private Sub cmdnext_Click()
frmequipmentallotment.Show

End Sub

Private Sub cmdsave_Click()
Adodc1.Recordset.Fields("ID") = txtid.Text
Adodc1.Recordset.Fields("NAME") = txteventselected.Text

Adodc1.Recordset.Fields("LOCATION") = txtlocation.Text
Adodc1.Recordset.Fields("OPERATION SELECTED") = txtoselected.Text
Adodc1.Recordset.Fields("COST OF EVENT") = txtcostofevent.Text
Adodc1.Recordset.Fields("COST OF OPERATION") = txtcostofoperation.Text
Adodc1.Recordset.Fields("TOTAL COST") = txttotalcost.Text




Adodc1.Recordset.Update
txteventselected.Text = cmbevent.Text
MsgBox ("Data Added Successfully")


End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub
