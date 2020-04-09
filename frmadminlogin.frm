VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmadminlogin 
   BackColor       =   &H00FF8080&
   Caption         =   "ADMIN LOGIN"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmadminlogin.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   14955
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   11760
      Top             =   4440
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   5280
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdadminlogin 
      BackColor       =   &H00C0C000&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8400
      MaskColor       =   &H00FFFF00&
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txtadminpassword 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtadminusername 
      Height          =   735
      Left            =   10200
      TabIndex        =   2
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lbladminpassword 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lbladminusername 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lbladminlogin 
      BackStyle       =   0  'Transparent
      Caption         =   "      ADMIN LOGIN"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmadminlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdadminlogin_Click()
ProgressBar1.Visible = True

Timer1.Enabled = True

If txtadminusername.Text = "user123" And txtadminpassword.Text = "officer123" Then



End If
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1

If ProgressBar1.Value = 100 Then
Timer1.Enabled = False

ProgressBar1.Visible = False
frmadminpoint.Show
End If
End Sub
