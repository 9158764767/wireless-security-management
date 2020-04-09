VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmuserlogin 
   BackColor       =   &H00FF8080&
   Caption         =   "USERLOGIN"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmuserlogin.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   17190
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1080
      Top             =   1200
   End
   Begin VB.CommandButton cmduserlogin 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9120
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txtpassword 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   11280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtusername 
      Height          =   855
      Left            =   11280
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lbluserlogin 
      BackStyle       =   0  'Transparent
      Caption         =   "USER LOGIN"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   5
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label lblpassword 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label lblusername 
      BackStyle       =   0  'Transparent
      Caption         =   " USERNAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "frmuserlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmduserlogin_Click()
ProgressBar1.Visible = True

Timer1.Enabled = True

If txtusername.Text = "user123" And txtpassword.Text = "officer123" Then


Else
MsgBox ("YOUR ACCESS IS NOT GRANTED AND YOU ARE NOT A GENUINE USER********")
End If


End Sub

Private Sub Timer1_Timer()



ProgressBar1.Value = ProgressBar1.Value + 1

If ProgressBar1.Value = 100 Then
Timer1.Enabled = False

ProgressBar1.Visible = False
frmuserpoint.Show
End If
End Sub

