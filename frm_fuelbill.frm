VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_fuelbill 
   Caption         =   "frm_fuelbill"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog printDialog1 
      Left            =   6120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label2 
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frm_fuelbill"
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
 Label1.Caption = frmfuelrequirement.txtrocketname
 Label2.Caption = frmfuelrequirement.txtfuelname
 Label3.Caption = frmfuelrequirement.txtfuelrequire
 Label4.Caption = frmfuelrequirement.txtfuelprice
End Sub
