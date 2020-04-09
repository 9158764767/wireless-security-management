VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_eventbill 
   Caption         =   "EVENT BILL"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   6120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   5760
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
      Height          =   735
      Left            =   5760
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label8 
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "TOTAL COST"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label6 
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "LOCATION"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label4 
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "OPERATION NAME"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "EVENT NAME"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frm_eventbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
printDialog1.ShowPrinter

End Sub

Private Sub cmdsave_Click()
printDialog1.ShowSave
End Sub

Private Sub Form_Load()
Label2.Caption = frmeventadnoperation.txteventselected.Text
Label4.Caption = frmeventadnoperation.txtoselected.Text
Label6.Caption = frmeventadnoperation.txtlocation.Text
Label8.Caption = frmeventadnoperation.txttotalcost.Text


End Sub
