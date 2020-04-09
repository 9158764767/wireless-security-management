VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_repairbill 
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   6960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   735
      Left            =   4920
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   4920
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label12 
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label11 
      Height          =   615
      Left            =   2400
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Height          =   735
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "REPAIRING PRICE"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "TOTAL LOSS"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "FAULT REPAIRED"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "EQUIPMENT PRICE"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "EQUIPMENT QUANTITY"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "EQUIPMENT NAME"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frm_repairbill"
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
Label7.Caption = frmtotallossandgainbyequipment.txteqname.Text
Label8.Caption = frmtotallossandgainbyequipment.txtquantity.Text
Label9.Caption = frmtotallossandgainbyequipment.txtprice.Text
Label10.Caption = frmtotallossandgainbyequipment.txtfaultresult.Text
Label11.Caption = frmtotallossandgainbyequipment.txtrepairprice.Text
Label12.Caption = frmtotallossandgainbyequipment.txtloss.Text

End Sub


