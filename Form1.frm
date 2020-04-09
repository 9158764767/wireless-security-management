VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmhomepage 
   BackColor       =   &H8000000D&
   Caption         =   "Homepage"
   ClientHeight    =   8550
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14070
   FillColor       =   &H00C0C000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1.15152e6
   ScaleMode       =   0  'User
   ScaleWidth      =   127.793
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   8055
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   873
      SimpleText      =   "DATE"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2534
            MinWidth        =   2534
            TextSave        =   "01-08-2017"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2538
            MinWidth        =   2538
            TextSave        =   "10:53 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Himalaya"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdsatellite 
      Caption         =   "WAY TO SATELLITE SECTION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   8880
      Picture         =   "Form1.frx":66C82
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton cmdcommunicate 
      Caption         =   "WAY TO COMMUNICATE SECTION"
      DownPicture     =   "Form1.frx":67A0C
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2400
      Picture         =   "Form1.frx":69823
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label lblhome 
      BackColor       =   &H8000000D&
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
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   13335
   End
   Begin VB.Menu mnuuser 
      Caption         =   "USER"
      Begin VB.Menu mnuequipmentrequirelist 
         Caption         =   "EQUIPMENT REQUIREMENT LIST"
      End
      Begin VB.Menu mnutypeofoperationlist 
         Caption         =   "TYPE OF EQUIPMENT AND OPERATION LIST"
      End
      Begin VB.Menu mnuequipmentinvoice 
         Caption         =   "OPERATION WISE PRICE OF EQUIPMENT(INVOICE)"
      End
      Begin VB.Menu mnunewpurchase 
         Caption         =   "NEW PURCHASE REQUIREMENT"
      End
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "ADMIN"
      Begin VB.Menu mnulogin 
         Caption         =   "LOGIN"
      End
      Begin VB.Menu mnuoffice 
         Caption         =   "NEW OFFICE"
      End
      Begin VB.Menu mnuemployee 
         Caption         =   "NEW EMPLOYEE"
      End
      Begin VB.Menu mnulossandgain 
         Caption         =   "TOTAL LOSS AND GAINS BY EQUIPMENT"
      End
      Begin VB.Menu mnureport 
         Caption         =   "REPORT"
      End
      Begin VB.Menu mnugovernment 
         Caption         =   "TOTAL INVOICE AND OUT STATEMENT FOR GOVERNMENT"
      End
   End
   Begin VB.Menu mnusatellite1 
      Caption         =   "SATELLITE"
      Index           =   1
      Begin VB.Menu mnusatelliteinfo 
         Caption         =   "SATELLITE INFORMATION"
      End
      Begin VB.Menu mnusatellite2 
         Caption         =   "LAUNCHING OF SATELLITE"
      End
      Begin VB.Menu mnutotalcost 
         Caption         =   "TOTAL COST OF SATELLITE"
      End
      Begin VB.Menu mnureducingphase 
         Caption         =   "REDUCING PHASE AND TOTAL COST"
      End
      Begin VB.Menu mnucostoffuel 
         Caption         =   "TOTAL COST OF FUEL/LAUNCHING OF SATELLITE"
      End
      Begin VB.Menu mnufrequency 
         Caption         =   "SETTING NEW FREQUENCY"
      End
      Begin VB.Menu mnuequipmentrequired 
         Caption         =   "EQUIPMENT REQUIRED FOR NEW SATELLITE"
      End
      Begin VB.Menu mnucompanywise 
         Caption         =   "COMPANY WISE PURCHASING COST OF SATELLITE EQUIPMENT"
      End
      Begin VB.Menu mnutotalinvoice 
         Caption         =   "TOTAL INVOICE "
      End
      Begin VB.Menu mnuexit 
         Caption         =   "EXIT"
      End
   End
End
Attribute VB_Name = "frmhomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
