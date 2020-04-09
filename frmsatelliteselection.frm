VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsatelliteselection 
   Caption         =   "SATELLITE SELECTION"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmsatelliteselection.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   18360
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   735
      Left            =   4680
      TabIndex        =   17
      Top             =   7680
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmsatelliteselection.frx":66C82
      Height          =   2415
      Left            =   840
      TabIndex        =   16
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtsatid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9240
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmsatelliteselection.frx":66C97
      OLEDBString     =   $"frmsatelliteselection.frx":66D2B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SATELLITE"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
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
      Left            =   10440
      TabIndex        =   13
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      TabIndex        =   11
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox txtfuelamt 
      DataField       =   "FUEL AMOUNT"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   9240
      TabIndex        =   10
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox txtsname 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   9240
      TabIndex        =   8
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "SELECT TYPE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   10695
      Begin VB.OptionButton optleo 
         BackColor       =   &H8000000D&
         Caption         =   "LEO"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7320
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optmeo 
         BackColor       =   &H8000000D&
         Caption         =   "MEO"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optgeo 
         BackColor       =   &H8000000D&
         Caption         =   "GEO"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   615
      Left            =   4560
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lblv 
      BackStyle       =   0  'Transparent
      Caption         =   "VERTICAL VELOCITY LAUNCH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   5160
      Width           =   5775
   End
   Begin VB.Label lblfuelamt 
      BackStyle       =   0  'Transparent
      Caption         =   "FUEL AMOUNT"
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
      Left            =   4320
      TabIndex        =   9
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label lblsname 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF SATELLITE"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label lblleo 
      BackStyle       =   0  'Transparent
      Caption         =   "LOWER EARTH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblmeo 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDIUM EARTH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblgeo 
      BackStyle       =   0  'Transparent
      Caption         =   "GEOSTATIONARY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
End
Attribute VB_Name = "frmsatelliteselection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
'If txtsatid.Text = "" Then
'MsgBox ("PLEASE ENTER THE FIELDS!!!!")
'End If
If txtsname.Text = "" Then
MsgBox ("PLEASE ENTER THE FIELDS!!!!")
End If
If txtfuelamt.Text = "" Then
MsgBox ("PLEASE ENTER THE FIELDS!!!!")
End If
Adodc1.Recordset.Fields("ID") = txtsatid.Text
Adodc1.Recordset.Fields("NAME") = txtsname.Text
Adodc1.Recordset.Fields("FUEL AMOUNT") = txtfuelamt.Text
frmsatelliteequipment.Show
End Sub

Private Sub cmdsave_Click()
 Adodc1.Recordset.Update
 
MsgBox ("SAVED SUCESSFULL!!!!!!!!!!!")


End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
txtfuelamt.Text = ""
txtsatid.Text = ""
txtsname.Text = ""
optgeo.Value = False
optleo.Value = False
optmeo.Value = False
End Sub

Private Sub optgeo_Click()
Adodc1.Recordset.Fields("ORBIT") = optgeo.Caption
txtfuelamt.Text = 2000000 & "metric ton"
End Sub

Private Sub optleo_Click()
Adodc1.Recordset.Fields("ORBIT") = optleo.Caption
     txtfuelamt.Text = 1000000 & "metric ton"
End Sub

Private Sub optmeo_Click()
Adodc1.Recordset.Fields("ORBIT") = optmeo.Caption
txtfuelamt.Text = 100000 & "metric ton"

End Sub
