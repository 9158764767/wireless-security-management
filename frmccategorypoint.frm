VERSION 5.00
Begin VB.Form frmccategorypoint 
   BackColor       =   &H00FF8080&
   Caption         =   "C CATEGORY SATELLITE GATEWAY"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   Picture         =   "frmccategorypoint.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "section"
      Height          =   1455
      Left            =   3840
      TabIndex        =   0
      Top             =   1440
      Width           =   3495
   End
End
Attribute VB_Name = "frmccategorypoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmccategorylogin.Show

End Sub
