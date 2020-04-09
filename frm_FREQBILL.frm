VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_FREQBILL 
   Caption         =   "frm_FREQBILL"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   7200
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label5 
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "TOTAL COST"
      Height          =   735
      Left            =   840
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frm_FREQBILL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsave_Click()
printDialog1.ShowSave
End Sub

Private Sub Command1_Click()
printDialog1.ShowPrinter
End Sub

Private Sub Form_Load()
Label1.Caption = frmnewfrequency.txtid.Text
Label2.Caption = frmnewfrequency.txtchannelset
Label3.Caption = frmnewfrequency.txtsetf
 Label5.Caption = frmnewfrequency.txtcost
End Sub

