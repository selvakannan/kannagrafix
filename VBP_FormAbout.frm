VERSION 5.00
Begin VB.Form FormAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About kannagrafix"
   ClientHeight    =   8490
   ClientLeft      =   2340
   ClientTop       =   1875
   ClientWidth     =   20745
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1383
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin kannagrafix.jcbutton jcbutton1 
      Height          =   375
      Index           =   0
      Left            =   10680
      TabIndex        =   2
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
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
      Caption         =   "Click on buttons below to open links;"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox lblThanks 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "kannagrafix developer site links;"
      Top             =   360
      Width           =   10455
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   7635
      TabIndex        =   0
      Top             =   7560
      Width           =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000002&
      X1              =   8
      X2              =   592
      Y1              =   16
      Y2              =   16
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'About Form
'Copyright ©2000-2012 by kannagrafix
'Created: 6/12/01
'Last updated: 04/September/12
'Last update: updated list to reflect recent changes to the codebase.
'
'A simple "about"/credits form.  Contains credits, copyright, and the program logo.
'
'***************************************************************************

Option Explicit
'ShellExecute is preferable to VB's 'Shell' command; I use it for two items in the "Help" menu - sending
' me an email, and opening the kannagrafix website (currently just tannerhelland.com)
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


Dim creditList() As String
Dim curCredit As Long

Private Sub CmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
   
    
    curCredit = 1
    
    'Shout-outs to other designers, programmers, testers and sponsors who provided various resources
    
    Create_Http_link "art4artist.com admin/art4artist/Karur@1234/.....http://art4artist.com/admin/", "http://art4artist.com/admin/"
    Create_Http_link "art4artist.com cpanel/art4artist/Karur@1234/.....http://art4artist.com/cpanel/", "http://art4artist.com/cpanel/"

    Create_Http_link "One note/kannanpainting@gmail.com/pass: karur1234...", "https://www.onenote.com/notebooks?wdorigin=ondcauth2%2condcnotebooks%2condchrd&auth=1&nf=1&fromAR=1"
    Create_Http_link "github kannagrafix / pass; karur123....https://github.com/selvakannan/kannagrafix ", "https://github.com/selvakannan/kannagrafix"
    Create_Http_link "Evernote../kannanpainting@gmail.com /..karur123../", "https://www.evernote.com/Home.action?_sourcePage=TlOCHGh2e8LiMUD9T65RG_YvRLZ-1eYO3fqfqRu0fynRL_1nukNa4gH1t86pc1SP&__fp=B7t-OXV5h6U3yWPvuidLz-TPR6I9Jhx8&hpts=1529553173298&showSwitchService=true&usernameImmutable=false&rememberMe=true&login=&login=Sign+in&login=true&username=kannanpainting%40gmail.com&hptsh=XoDpMukOrj0G6o9tOKLAg2Kd3UE%3D#n=86eab3ea-a263-4c1f-bf09-d2850c9a36ce&s=s397&ses=4&sh=2&sds=5&"
    Create_Http_link "freenom", "http://www.freenom.com/en/index.html?lang=en"
    Create_Http_link "Blogger/ blogsppot", "https://www.blogger.com/blogger.g?blogID=7552986963327589478#editor/target=post;postID=4357116948011763365;onPublishedMenu=posts;onClosedMenu=posts;postNum=0;src=link"
    Create_Http_link "planet-source-code.com/kannanpainting@gmail.com/karur1234/", "http://www.planet-source-code.com/vb/default.asp"

   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   
    Create_Http_link "by selvaraj kannan"
   
    
    lblThanks(0).MousePointer = vbDefault
    
End Sub

'Generate a label with the specified "thank you" text, and link it to the specified URL
Private Sub Create_Http_link(ByVal thxText As String, Optional ByVal creditURL As String = "")
    
    'Generate a new label
    Load lblThanks(curCredit)
    Load jcbutton1(curCredit)
    'Because I now have too many people to thank, it's necessary to split the list into two columns
    Dim columnLimit As Long
    columnLimit = 19
    
    If curCredit = 1 Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 12
        lblThanks(curCredit).Left = lblThanks(0).Left + 2
        jcbutton1(curCredit).Top = jcbutton1(curCredit - 1).Top + jcbutton1(curCredit - 1).Height + 12
        jcbutton1(curCredit).Left = jcbutton1(0).Left + 2
    ElseIf curCredit < columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 2
        jcbutton1(curCredit).Top = jcbutton1(curCredit - 1).Top + jcbutton1(curCredit - 1).Height + 4
        jcbutton1(curCredit).Left = jcbutton1(0).Left + 2
    ElseIf curCredit = columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 12 - (lblThanks(columnLimit - 1).Top - lblThanks(0).Top)
        lblThanks(curCredit).Left = lblThanks(0).Left + 325
        jcbutton1(curCredit).Top = jcbutton1(curCredit - 1).Top + jcbutton1(curCredit - 1).Height + 12 - (jcbutton1(columnLimit - 1).Top - jcbutton1(0).Top)
        jcbutton1(curCredit).Left = jcbutton1(0).Left + 325
    Else
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 325
         jcbutton1(curCredit).Top = jcbutton1(curCredit - 1).Top + jcbutton1(curCredit - 1).Height + 4
        jcbutton1(curCredit).Left = jcbutton1(0).Left + 325
    End If
    lblThanks(curCredit).FontName = "tahoma"
    lblThanks(curCredit).FontBold = True
    lblThanks(curCredit).Text = thxText
    jcbutton1(curCredit).Caption = thxText

    If creditURL = "" Then
        lblThanks(curCredit).MousePointer = vbDefault
        jcbutton1(curCredit).MousePointer = vbDefault
    Else
        lblThanks(curCredit).FontUnderline = True
        lblThanks(curCredit).ForeColor = vbBlue
        lblThanks(curCredit).ToolTipText = "Click to open " & creditURL
                jcbutton1(curCredit).ToolTip = "Click to open " & creditURL

    End If
    lblThanks(curCredit).Visible = True
        jcbutton1(curCredit).Visible = True

    ReDim Preserve creditList(0 To curCredit) As String
    creditList(curCredit) = creditURL
    
    curCredit = curCredit + 1

End Sub


Private Sub jcbutton1_Click(Index As Integer)
    If creditList(Index) <> "" Then ShellExecute FormMain.HWnd, "Open", creditList(Index), "", 0, SW_SHOWNORMAL

End Sub

Private Sub lblThanks_DblClick(Index As Integer)
    If creditList(Index) <> "" Then ShellExecute FormMain.HWnd, "Open", creditList(Index), "", 0, SW_SHOWNORMAL

End Sub
