VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_newpurchasebill 
   Caption         =   "NEW PURCHASE BILL"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   6240
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   1215
      Left            =   6120
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label8 
      Height          =   1215
      Left            =   2760
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "TOTAL PRICE"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label6 
      Height          =   1095
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "COMPANY NAME"
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label4 
      Height          =   975
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "EQUIPMENT NAME"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frm_newpurchasebill"
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
Label2.Caption = frmnewpurchase.txtid.Text
Label4.Caption = frmnewpurchase.txtename.Text

Label6.Caption = frmnewpurchase.txtcompany.Text
Label8.Caption = frmnewpurchase.txteqprice.Text


End Sub
