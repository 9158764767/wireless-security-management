VERSION 5.00
Begin VB.Form frmuserpoint 
   BackColor       =   &H00FF8080&
   Caption         =   "USER GATEWAY"
   ClientHeight    =   8685
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmuserpoint.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   15555
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuuser 
      Caption         =   "USER"
      Begin VB.Menu mnuequipmentrequirementlist 
         Caption         =   "EQUIPMENT REQUIREMENT LIST"
      End
      Begin VB.Menu mnuequipmentuseinoperation 
         Caption         =   "EQUIPMENT USE IN OPERATION"
      End
      Begin VB.Menu mnunewpurchase 
         Caption         =   "NEW PURCHASE"
      End
      Begin VB.Menu mnueventandoperation 
         Caption         =   "EVENTS AND OPERATION"
      End
      Begin VB.Menu mnuequipmentallotment 
         Caption         =   "EQUIPMENT ALLOTMENT"
      End
      Begin VB.Menu mnurepairsection 
         Caption         =   "REPAIRING SECTION"
      End
      Begin VB.Menu mnueport 
         Caption         =   "REPORT"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "frmuserpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcommunicate_Click()
frmuserlogin.Show

End Sub



Private Sub mnueport_Click()
frmgoreport.Show

End Sub

Private Sub mnuequipmentallotment_Click()
frmequipmentallotment.Show

End Sub

Private Sub mnuequipmentrequirementlist_Click()
frmequipmentrequirement.Show

End Sub

Private Sub mnuequipmentuseinoperation_Click()
frmequipmentusesinoperation.Show

End Sub

Private Sub mnueventandoperation_Click()
frmeventadnoperation.Show

End Sub

Private Sub mnunewpurchase_Click()
frmnewpurchase.Show

End Sub

Private Sub mnurepairsection_Click()
frmrepairingsection.Show

End Sub

Private Sub mnutypeofoperation_Click()
frmequipmentusesinoperation.Show

End Sub

Private Sub mnuuserlogin_Click()
frmuserlogin.Show

End Sub
