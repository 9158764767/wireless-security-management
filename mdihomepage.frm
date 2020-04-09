VERSION 5.00
Begin VB.MDIForm mdihomepage 
   BackColor       =   &H8000000C&
   Caption         =   "HOMEPAGE"
   ClientHeight    =   7980
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14010
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdihomepage.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuuser 
      Caption         =   "USER"
      Visible         =   0   'False
      Begin VB.Menu mnuuserlogin 
         Caption         =   "USER LOGIN"
      End
      Begin VB.Menu mnuequipmentrequirement 
         Caption         =   "EQUIPMENT REQUIRMENT"
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
         Caption         =   "REPAIR SECTION"
      End
      Begin VB.Menu mnureport1 
         Caption         =   "REPORT"
      End
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "ADMIN"
      Visible         =   0   'False
      Begin VB.Menu mnuadminlogin 
         Caption         =   "ADMIN LOGIN"
      End
      Begin VB.Menu mnunewoffice 
         Caption         =   "NEW OFFICE"
      End
      Begin VB.Menu mnunewemployee 
         Caption         =   "NEW EMPLOYEE"
      End
      Begin VB.Menu mnutotallossandgain 
         Caption         =   "TOTAL LOSS AND GAIN"
      End
      Begin VB.Menu mnugovernment 
         Caption         =   "TOTAL INVOICE AND OUT STATEMENT FOR GOVERNMENT"
      End
      Begin VB.Menu mnureport2 
         Caption         =   "REPORT"
      End
   End
   Begin VB.Menu mnusatellite 
      Caption         =   "SATELLITE"
      Visible         =   0   'False
      Begin VB.Menu mnusatelliteselection 
         Caption         =   "SATELLITE SELECTION"
      End
      Begin VB.Menu mnufuelrequirement 
         Caption         =   "FUEL REQUIREMENT"
      End
      Begin VB.Menu mnurocketselection 
         Caption         =   "ROCKET SELECTION"
      End
      Begin VB.Menu mnusatellitelaunching 
         Caption         =   "SATELLITE LAUNCHING"
      End
      Begin VB.Menu mnutotalcost 
         Caption         =   "TOTAL COST"
      End
      Begin VB.Menu mnufrequency 
         Caption         =   "SETTING NEW FREQUENCY"
      End
      Begin VB.Menu mnuequipmentrequired 
         Caption         =   "EQUIPMENT REQUIRED"
      End
      Begin VB.Menu mnureport3 
         Caption         =   "TOTAL REPORT"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "EXIT"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdihomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuadminlogin1_Click()
frmadminlogin.Show

End Sub

Private Sub mnuadminlogin2_Click()
frmadminlogin.Show

End Sub

Private Sub mnuadminlogin_Click()
frmadminlogin.Show

End Sub

Private Sub mnuequipmentallotment_Click()
frmequipmentallotment.Show

End Sub

Private Sub mnuequipmentrequired_Click()
frmsatelliteequipment.Show

End Sub

Private Sub mnuequipmentrequirement_Click()
frmequipmentrequirement.Show

End Sub

Private Sub mnuequipmentuseinoperation_Click()
frmequipmentusesinoperation.Show

End Sub

Private Sub mnueventandoperation_Click()
frmeventadnoperation.Show

End Sub

Private Sub mnuexit_Click()
End

End Sub

Private Sub mnufrequency_Click()
frmnewfrequency.Show

End Sub

Private Sub mnufuelrequirement_Click()
frmfuelrequirement.Show

End Sub

Private Sub mnunewemployee_Click()
frmnewemployee.Show

End Sub

Private Sub mnunewoffice_Click()
frmnewoffice.Show

End Sub

Private Sub mnunewpurchase_Click()
frmnewpurchase.Show

End Sub

Private Sub mnurepairsection_Click()
frmrepairingsection.Show

End Sub

Private Sub mnurocketselection_Click()
frmrocketselection.Show

End Sub

Private Sub mnusatellitelaunching_Click()
frmlaunching.Show

End Sub

Private Sub mnusatelliteselection_Click()
frmsatelliteselection.Show

End Sub

Private Sub mnutotalcost_Click()
frmtotalcostofsatellite.Show

End Sub

Private Sub mnutotallossandgain_Click()
frmtotallossandgainbyequipment.Show

End Sub

Private Sub mnuuserlogin_Click()
frmuserlogin.Show

End Sub
