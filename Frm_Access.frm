VERSION 5.00
Begin VB.Form Frm_VB6_Access 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmdConnectMySQL 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin kannagrafix.jcbutton jcbutton2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "student info"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin kannagrafix.jcbutton jcbutton1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "stock"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
End
Attribute VB_Name = "Frm_VB6_Access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Visual Basic 6 code for connecting to aMySQL database using the MySQL ODBCdriver.
'This code snippet demonstrates how to connect to a MySQL database from a Windows basedapplication written in Visual Basic 6. By using the MySQL ODBC driver and the MicrosoftRemote Data Objectit is quite easy to connect and retrieve records from a MySQL databaseserver.
'•
'Download and install theMySQL ODBC driver .
'•
'Set-up a MySQL username and password combination that will allow connections fromany host. See MySQLs
'grant
'command.
'•
'Start a new Visual Basic project and add the Microsoft Remote Data Object - Using themenus select Project | References and then select the Microsoft Remote Data Object fromthe list.
'Sample Code

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim username As String
Dim passwd As String
Dim serverIP As String
Dim db As String
Dim ssql As String
Private Sub cmdConnectMySQL_Click()
Dim cnMySql As New rdoConnection
Dim rdoQry As New rdoQuery
Dim rdoRS As rdoResultset
' set up a remote data connection
' using the MySQL ODBC driver.
' change the connect string with your username,
' password, server name and the database you
' wish to connect to.
cnMySql.CursorDriver = rdUseOdbc
cnMySql.Connect = "uid=YourUserName;pwd=YourPassword;server=YourServerName;" & "driver={MySQL ODBC 3.51 Driver};database=YourDataBase;dsn=;"
cnMySql.EstablishConnection
' set up a remote data object query
' specifying the SQL statement to run.
With rdoQry
.Name = "selectUsers"
.SQL = "select * from user"
.RowsetSize = 1
Set .ActiveConnection = cnMySql
Set rdoRS = .OpenResultSet(rdOpenKeyset, rdConcurRowVer)
End With

End Sub


Private Sub Command1_Click()
' put this in your form

Call connectMysql("root", "", "127.0.0.1", "myerp", conn, rs)

ssql = "SELECT * FROM accmaster"
rs.Open ssql, conn

Set DataGrid1.DataSource = rs

End Sub

Private Sub jcbutton1_Click()
dblWordID = Shell("vb6/access/stock/Stock.EXE", vbNormalFocus)
     AppActivate dblWordID

End Sub

Private Sub jcbutton2_Click()
 dblWordID = Shell("vb6/access/DATA-KG/StdInfo.EXE", vbNormalFocus)
     AppActivate dblWordID
'Data -KG

End Sub
