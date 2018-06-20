VERSION 5.00
Begin VB.Form FormAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About kannagrafix"
   ClientHeight    =   8115
   ClientLeft      =   2340
   ClientTop       =   1875
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
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
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Label lblThanks 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "kannagrafix would not be possible without the following individuals. My sincerest thanks goes out to: "
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   7275
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Copyright (automatically populated at run-time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2910
      TabIndex        =   2
      Top             =   2400
      Width           =   5985
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version (automatically populated at run-time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   3900
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
    
    'Automatic generation of version & copyright information
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDisclaimer.Caption = App.LegalCopyright & "   "
    
    curCredit = 1
    
    'Shout-outs to other designers, programmers, testers and sponsors who provided various resources
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
    Create_Http_link "Kannagrafix founder of art4artist.com website", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   Create_Http_link "Kannagrafix", "http://art4artist.com"
   
    Create_Http_link "selvaraj kannan"
    Create_Http_link "ponni kannan"
    Create_Http_link "santhish kannan"
    
    lblThanks(0).MousePointer = vbDefault
    
End Sub

'Generate a label with the specified "thank you" text, and link it to the specified URL
Private Sub Create_Http_link(ByVal thxText As String, Optional ByVal creditURL As String = "")
    
    'Generate a new label
    Load lblThanks(curCredit)
    
    'Because I now have too many people to thank, it's necessary to split the list into two columns
    Dim columnLimit As Long
    columnLimit = 19
    
    If curCredit = 1 Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 12
        lblThanks(curCredit).Left = lblThanks(0).Left + 2
    ElseIf curCredit < columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 2
    ElseIf curCredit = columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 12 - (lblThanks(columnLimit - 1).Top - lblThanks(0).Top)
        lblThanks(curCredit).Left = lblThanks(0).Left + 325
    Else
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 325
    End If
    
    lblThanks(curCredit).Caption = thxText
    If creditURL = "" Then
        lblThanks(curCredit).MousePointer = vbDefault
    Else
        lblThanks(curCredit).FontUnderline = True
        lblThanks(curCredit).ForeColor = vbBlue
        lblThanks(curCredit).ToolTipText = "Click to open " & creditURL
    End If
    lblThanks(curCredit).Visible = True
    
    ReDim Preserve creditList(0 To curCredit) As String
    creditList(curCredit) = creditURL
    
    curCredit = curCredit + 1

End Sub

'When a thank-you credit is clicked, launch the corresponding website
Private Sub lblThanks_Click(Index As Integer)

    If creditList(Index) <> "" Then ShellExecute FormMain.HWnd, "Open", creditList(Index), "", 0, SW_SHOWNORMAL

End Sub
