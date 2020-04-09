VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmgateway 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "GATEWAY"
   ClientHeight    =   8415
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   16275
   FillColor       =   &H00C0C000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "frmhomepage.frx":0000
   ScaleHeight     =   1.13334e6
   ScaleMode       =   0  'User
   ScaleWidth      =   147.82
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   1095
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdgateway 
      Caption         =   "GATEWAY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3480
      TabIndex        =   6
      Top             =   2160
      Width           =   9975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8040
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2536
            MinWidth        =   2536
            TextSave        =   "24-09-2017"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2538
            MinWidth        =   2538
            TextSave        =   "12:06 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "4)  This system is used in signalling department in army in filt tactics called as Field Martial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   6840
      Width           =   15615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3)  Department for communication between variousoperations and digital connectivity."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   14295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2)  Recently this system is used in State Police "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   5400
      Width           =   13215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1)  This system is used in various defence system in Indian Government and worldwide."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   14775
   End
   Begin VB.Label lblhome 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "WIRELESS SECURITY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmgateway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcommunicate_Click()
frmequipmentrequirement.Show

End Sub

Private Sub cmdsatellite_Click()
frmadminlogin.Show

End Sub

Private Sub mnuclogin_Click()
frmccategorypoint.Show

End Sub

Private Sub mnucostoffuel_Click()
frmfuelrequirement.Show

End Sub

Private Sub mnuemployee_Click()
frmnewemployee.Show

End Sub

Private Sub mnuequipmentallotment_Click()
frmequipmentallotment.Show

End Sub

Private Sub mnuequipmentrequired_Click()
frmsatelliteequipment.Show

End Sub

Private Sub mnuequipmentrequirelist_Click()
frmequipmentrequirement.Show
End Sub

Private Sub mnueventsandoperations_Click()
frmeventadnoperation.Show

End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnufrequency_Click()
frmnewfrequency.Show

End Sub

Private Sub mnulogin_Click()
frmadminpoint.Show


End Sub



Private Sub mnulossandgain_Click()
frmtotallossandgainbyequipment.Show

End Sub

Private Sub mnunewpurchase_Click()
frmnewpurchase.Show

End Sub

Private Sub mnuoffice_Click()
frmnewoffice.Show

End Sub

Private Sub mnurepairsection_Click()
frmrepairingsection.Show

End Sub

Private Sub mnurocketselection_Click()
frmrocketselection.Show

End Sub

Private Sub mnusatellitelaunch_Click()
frmlaunching.Show

End Sub

Private Sub mnusatelliteselect_Click()
frmsatelliteselection.Show

End Sub

Private Sub mnutotalcost_Click()
frmtotalcostofsatellite.Show

End Sub

Private Sub mnutypeofoperationlist_Click()
frmequipmentusesinoperation.Show

End Sub

Private Sub mnuuserlogin_Click()
frmuserpoint.Show


End Sub

Private Sub cmdgateway_Click()
frmabcd.Show



End Sub

Private Sub Command1_Click()
End
End Sub
