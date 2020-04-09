VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmnewemployee 
   BackColor       =   &H00FF8080&
   Caption         =   "NEW EMPLOYEE REGISTRATION"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15495
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmnewemployee.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   15495
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      Caption         =   "female"
      Height          =   375
      Left            =   11640
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "male"
      Height          =   495
      Left            =   9960
      TabIndex        =   30
      Top             =   2880
      Width           =   975
   End
   Begin MSACAL.Calendar Calendar3 
      Height          =   2055
      Left            =   2880
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   3855
      _Version        =   524288
      _ExtentX        =   6800
      _ExtentY        =   3625
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2017
      Month           =   9
      Day             =   9
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Calendar2 
      Height          =   2175
      Left            =   9600
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   3836
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2017
      Month           =   9
      Day             =   9
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   9960
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2017
      Month           =   9
      Day             =   9
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtservice 
      DataField       =   "SERVICE"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   10200
      TabIndex        =   26
      Top             =   6360
      Width           =   3015
   End
   Begin VB.TextBox txtsalary 
      DataField       =   "SALARY"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   25
      Top             =   5280
      Width           =   3855
   End
   Begin VB.TextBox txtprofile 
      DataField       =   "DESIGNITION"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   2640
      TabIndex        =   22
      Top             =   3840
      Width           =   4335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmnewemployee.frx":66C82
      Height          =   1215
      Left            =   11640
      TabIndex        =   21
      Top             =   7680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2143
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9360
      Top             =   8040
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1296
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
      Connect         =   $"frmnewemployee.frx":66C97
      OLEDBString     =   $"frmnewemployee.frx":66D2B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "EMPLOYEE"
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
      Height          =   855
      Left            =   3960
      TabIndex        =   20
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   18
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdregister 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   17
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox txtretiredate 
      DataField       =   "DATE OF RETIREMENT"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   2640
      TabIndex        =   16
      Top             =   5160
      Width           =   4215
   End
   Begin VB.TextBox txtjdate 
      DataField       =   "DATE OF JOINING"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   14
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox txtgender 
      DataField       =   "GENDER"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10080
      TabIndex        =   11
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox txTallotedsection 
      DataField       =   "ALLOTED SECTION NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2640
      TabIndex        =   9
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox txtdob 
      DataField       =   "DATE OF BIRTH"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   7
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   10080
      TabIndex        =   5
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtrecordno 
      DataField       =   "RECORD NO"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtname 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblservice 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVICE"
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
      Left            =   7680
      TabIndex        =   24
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label lblsalary 
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   23
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblretiredate 
      BackStyle       =   0  'Transparent
      Caption         =   "RETIREMENT DATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label lbljdate 
      BackStyle       =   0  'Transparent
      Caption         =   "JOINING DATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   13
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblprofile 
      BackStyle       =   0  'Transparent
      Caption         =   "PROFILE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblgender 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   10
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblcontact 
      BackStyle       =   0  'Transparent
      Caption         =   "ALLOTED SECTION"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lbldob 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblid 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblrecordno 
      BackStyle       =   0  'Transparent
      Caption         =   "RECORD NO"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblname 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmnewemployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
txtdob.Text = Calendar1.Value
Calendar1.Visible = False

End Sub

Private Sub Calendar2_Click()
txtjdate.Text = Calendar2.Value
Calendar2.Visible = False

End Sub

Private Sub Calendar3_Click()
txtretiredate.Text = Calendar3.Value
Calendar3.Visible = False

End Sub

Private Sub cmdnext_Click()
frmtotallossandgainbyequipment.Show

End Sub

Private Sub cmdregister_Click()
Adodc1.Recordset.Fields("ID") = txtid.Text
Adodc1.Recordset.Fields("NAME") = txtname.Text
Adodc1.Recordset.Fields("GENDER") = txtgender.Text
Adodc1.Recordset.Fields("DATE OF BIRTH") = txtdob.Text
Adodc1.Recordset.Fields("DATE OF JOINING") = txtjdate.Text
Adodc1.Recordset.Fields("DATE OF RETIREMENT") = txtretiredate.Text
Adodc1.Recordset.Fields("SERVICE") = txtservice.Text
Adodc1.Recordset.Fields("RECORD NO") = txtrecordno.Text
Adodc1.Recordset.Fields("DESIGNITION") = txtprofile.Text
Adodc1.Recordset.Fields("ALLOTED SECTION NAME") = txTallotedsection.Text
Adodc1.Recordset.Fields("SALARY") = txtsalary.Text

Adodc1.Recordset.Update
MsgBox ("data added successfully")







End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub

Private Sub Option1_Click()
txtgender.Text = Option1.Caption


End Sub

Private Sub Option2_Click()
txtgender.Text = Option2.Caption

End Sub

Private Sub txtdob_click()
Calendar1.Visible = True



End Sub

Private Sub txtjdate_click()
Calendar2.Visible = True

End Sub

Private Sub txtretiredate_Click()
Calendar3.Visible = True

End Sub
