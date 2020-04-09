VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_satbill 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton save 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   4920
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   8760
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "BILL PRINT"
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label lbltot 
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "TOTAL PRICE"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lbllp 
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "LOUNCHING PRICE"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblrc 
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "ROCKET PRICE"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblsp 
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "SATELLITE PRICE"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frm_satbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
'PrintDialog printDialog1 = new PrintDialog()
'printDialog1.Document = printDocument1
'DialogResult result = printDialog1.ShowDialog(this)
'printDialog1.Print()
printDialog1.ShowPrinter

End Sub

Private Sub Form_Load()
lblsp.Caption = frmtotalcostofsatellite.txtsatprice.Text
lblrc.Caption = frmtotalcostofsatellite.txtrocketcost.Text
lbllp.Caption = frmtotalcostofsatellite.txtlaunchcost.Text
lbltot.Caption = frmtotalcostofsatellite.txtscost.Text

End Sub

Private Sub save_Click()
printDialog1.ShowSave

End Sub
