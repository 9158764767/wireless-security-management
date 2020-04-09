VERSION 5.00
Begin VB.Form frmccategorylogin 
   BackColor       =   &H00FF8080&
   Caption         =   "C CATEGORY LOGIN"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   Picture         =   "frmccategorylogin.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3240
      TabIndex        =   4
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   5040
      TabIndex        =   1
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "          C CATEGORY LOGIN"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
   End
End
Attribute VB_Name = "frmccategorylogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
