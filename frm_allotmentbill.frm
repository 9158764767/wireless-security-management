VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_allotmentbill 
   Caption         =   "ALLOTMENT BILL"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   6120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   735
      Left            =   5880
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   735
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Height          =   975
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "QUANTITY"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label4 
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "EQUIPMENT NAME"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "EMPLOYEE NAME"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frm_allotmentbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
printDialog1.ShowPrinter
End Sub

Private Sub Command2_Click()
printDialog1.ShowSave
End Sub

Private Sub Form_Load()
Label2.Caption = frmequipmentallotment.txtempname.Text
Label4.Caption = frmequipmentallotment.txtselectequipment.Text

Label6.Caption = frmequipmentallotment.txteququantity.Text

'Text2.Text = Date


End Sub
