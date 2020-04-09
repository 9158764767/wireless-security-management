VERSION 5.00
Begin VB.Form frmadminpoint 
   BackColor       =   &H00FF8080&
   Caption         =   "ADMIN GATEWAY"
   ClientHeight    =   8235
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmadminpoint.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   15735
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuadmin 
      Caption         =   "ADMIN"
      Begin VB.Menu mnunewoffice 
         Caption         =   "NEW OFFICE"
      End
      Begin VB.Menu mnunewemployee 
         Caption         =   "NEW EMPLOYEE"
      End
      Begin VB.Menu mnutotallossandgain 
         Caption         =   "TOTAL LOSS AND GAIN"
      End
      Begin VB.Menu mnureport 
         Caption         =   "REPORT"
      End
   End
End
Attribute VB_Name = "frmadminpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnunewemployee_Click()
frmnewemployee.Show

End Sub

Private Sub mnunewoffice_Click()
frmnewoffice.Show

End Sub

Private Sub mnureport_Click()
frmgoreport.Show

End Sub

Private Sub mnutotallossandgain_Click()
frmtotallossandgainbyequipment.Show

End Sub
