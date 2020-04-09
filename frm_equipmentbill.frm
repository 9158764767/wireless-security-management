VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_equipmentbill 
   Caption         =   "EQUIPMENT BILL"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   5760
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   735
      Left            =   5400
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "print"
      Height          =   855
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "TOTAL COST"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Height          =   855
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "OPERATION NAME"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Caption         =   "EQUIPMENT NAME"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frm_equipmentbill"
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
 Label2.Caption = frmequipmentusesinoperation.txtequipmentselected.Text
 
 Label4.Caption = frmequipmentusesinoperation.txtoselected.Text
 
 Label6.Caption = frmequipmentusesinoperation.txttotalcost.Text
 

End Sub
