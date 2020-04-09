VERSION 5.00
Begin VB.Form frmabcd 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "satellite"
      Height          =   1515
      Left            =   6360
      TabIndex        =   2
      Top             =   4920
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "admin"
      Height          =   1575
      Left            =   6480
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "user"
      Height          =   1455
      Left            =   6480
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmabcd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmuserlogin.Show

End Sub

Private Sub Command2_Click()
frmadminlogin.Show

End Sub

Private Sub Command3_Click()
MsgBox ("THIS IS CONFEDENTIAL ACCESS!!!!!!!!!")
frmsatellitepoint.Show


End Sub
