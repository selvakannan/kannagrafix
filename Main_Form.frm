VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Email (SMTP) Client For Gmail"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   14415
   StartUpPosition =   2  'CenterScreen
   Begin Desktop_Email_Client.jcbutton jcbutton1 
      Height          =   615
      Left            =   11040
      TabIndex        =   31
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8640
      TabIndex        =   29
      Text            =   "raman   kannan  Karur@1234"
      Top             =   5160
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8640
      TabIndex        =   28
      Text            =   "kannanpainting@gmail.com"
      Top             =   5760
      Width           =   5055
   End
   Begin VB.ListBox ListBox1 
      Height          =   1230
      Left            =   360
      TabIndex        =   13
      Top             =   4680
      Width           =   6855
   End
   Begin VB.TextBox TextBox1 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Main_Form.frx":0000
      Top             =   6240
      Width           =   7575
   End
   Begin VB.TextBox TextBox9 
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox TextBox4 
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox TextBox2 
      Height          =   615
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox TextBox10 
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.TextBox TextBox5 
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   1320
      Width           =   5655
   End
   Begin VB.TextBox TextBox3 
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox TextBox6 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   11535
   End
   Begin VB.TextBox TextBox7 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2760
      Width           =   11415
   End
   Begin VB.TextBox TextBox8 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   11415
   End
   Begin Desktop_Email_Client.jcbutton CommandButton3 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      buttonstyle     =   2
      font            =   "Main_Form.frx":0005
      backcolor       =   15199212
      caption         =   "Delete"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin Desktop_Email_Client.jcbutton CommandButton2 
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      buttonstyle     =   2
      font            =   "Main_Form.frx":002D
      backcolor       =   15199212
      caption         =   "Attach Files"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin Desktop_Email_Client.jcbutton CommandButton1 
      Height          =   615
      Left            =   8400
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
      _extentx        =   3836
      _extenty        =   1085
      buttonstyle     =   2
      font            =   "Main_Form.frx":0055
      backcolor       =   15199212
      caption         =   "Send Email"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Desktop_Email_Client.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   8520
      TabIndex        =   30
      Top             =   4800
      Width           =   5535
      _extentx        =   9763
      _extenty        =   661
   End
   Begin VB.Label Label14 
      Caption         =   "Red colored are important."
      Height          =   255
      Left            =   6240
      TabIndex        =   27
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "bcc"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "cc"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   9240
      TabIndex        =   23
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Receiver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "EMAIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7200
      TabIndex        =   21
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "EMAIL"
      Height          =   255
      Left            =   7680
      TabIndex        =   20
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Sender"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Mail id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7320
      TabIndex        =   17
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "465"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "smtp.gmail.com"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Label14"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   6135
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   8040
      X2              =   0
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   8040
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   8040
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command2_Click()
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'On some requests when same project uploaded on PSC with  '
'the name of 'Urdu Desktop (SMTP) Email Client for Gmail' '
'Some users required to get the english version of the    '
'same project then it is converted by me to english &     '
'being uploaded on Planet Source Code.                    '
'There is a desktop email client to send email using gmail'
'address. No need to go on Gmail's website & login to send'
'emails. Simply put you Gmail address & Password & send   '
'email with the (n) number of atachments. No size limit to'
'attach files with email. All files with large size will  '
'be easily sent by this desktop client. CC & BCC function '
'also support you to send emails with large attachments,  '
'to (n) number of receivers. Gmail's SMTP address & port  '
'fixed in this client. Keep in mind that this email client'
'only designed to work with Gmail. No other email service '
'provider checked with this email client and that may take'
'errors, if you try to do that. This project may need more'
'attention but at the start, i think it is enough to use. '
'                                                         '
' Waiting for your Feedbacks.Thank You.                   '
'                                                         '
'                                                         '
'                 Muhammad Mehmood Iqbal                  '
'                   ME_IQ_TM@yahoo.com                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CommandButton1_Click()
SendEmail TextBox5, TextBox3, TextBox8, "from" & TextBox4.Text & "Time" & Now & "--" & TextBox1
End Sub

  

Private Sub CommandButton2_Click()

  Dim File_Name
  
  'Set Dialogbox Title
  CommonDialog1.DialogTitle = "Select File to Attach"

  ' Set flags
  CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
  
  ' Set filters
  CommonDialog1.Filter = "All Files (*.*)"
  
  ' Specify default filter
  CommonDialog1.FilterIndex = 1
  
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  
  If CommonDialog1.CancelError = True Or CommonDialog1.FileName = "" Then
  GoTo Exit_Sub
  Else
  
  'Count attached file show them
  ListBox1.AddItem CommonDialog1.FileName
  Label9.Caption = Main_Form.ListBox1.ListCount & " File (s) attached at this time."
  
  'Enable delete button
  CommandButton3.Enabled = True
  
  CommonDialog1.FileName = ""
  
  End If
   
Exit_Sub:
End Sub

Private Sub CommandButton3_Click()

Dim Selected_Item

'Check if no-item seleted in listbox
ListBox1.SetFocus
If ListBox1.ListIndex = -1 Then
GoTo End_Sub:

ElseIf ListBox1.ListIndex >= 0 Then

'Delete selected item
Selected_Item = ListBox1.ListIndex
ListBox1.RemoveItem (Selected_Item)

'Count -1 from attached files
Label9.Caption = Main_Form.ListBox1.ListCount & " File (s) attached at this time."

End If

'If no file in Listbox then disable Delete Button
If ListBox1.ListCount = 0 Then
CommandButton3.Enabled = False
ListBox1.SetFocus

End If

End_Sub:
End Sub

Private Sub Form_Load()


'Disable delete button of attached file
CommandButton3.Enabled = False

'Set attached file's status
Label9.Caption = Main_Form.ListBox1.ListCount

End Sub
' From http://www.vbknowledgebase.com/?Id=21&Desc=Send-Email(E-Mail)-from-VB6-using-CDO

'****************************************************************
'*  Purpose :   To Send eMail
'*
'*  Inputs  :   strRecipient(String)    Recipient comma seperated
'*              strSubject(String)      Subject
'*              strBody                  Body
'*              colAttachments          Collection of attachments
'*                                      file paths.
'*
'*  Returns :   Boolean about the sent status
'****************************************************************
Public Function SendEmail(ByVal strSender As String, _
                        ByVal strRecipient As String, _
                        ByVal strSubject As String, _
                        ByVal strBody As String, _
                        Optional ByVal strCc As String, _
                        Optional ByVal strBcc As String, _
                        Optional ByVal colAttachments As Collection _
                         ) As Boolean
    Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim schema As String
    Dim Flds
    Dim attachment
    Dim strHTML
    
    On Error GoTo ErrTrap
    Const cdoSendUsingPort = 2
    
    'Set cdoMsg =  CreateObject("CDO.Message")
    'Set cdoConf = CreateObject("CDO.Configuration")
    
    Set Flds = cdoConf.Fields
        
    schema = "http://schemas.microsoft.com/cdo/configuration/"

    With Flds
        .Item(schema & "sendusing") = 2
        .Item(schema & "smtpserver") = "smtp.gmail.com"
        .Item(schema & "smtpserverport") = 465
        .Item(schema & "smtpauthenticate") = 1
        .Item(schema & "sendusername") = TextBox5.Text
        .Item(schema & "sendpassword") = TextBox10.Text
        .Item(schema & "smtpusessl") = 1
        .Update
    End With
    XP_ProgressBar1.Value = 30
    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = strRecipient
        .From = strSender
        .Subject = strSubject
        .TextBody = strBody
        If Not colAttachments Is Nothing Then
            For Each attachment In colAttachments
                .AddAttachment attachment
            Next
        End If
        If strCc <> "" Then .CC = Main_Form.TextBox6.Text
        If strBcc <> "" Then .Bcc = Main_Form.TextBox7.Text
        'Check If Files attached then send them one by one
For Index = 0 To Main_Form.ListBox1.ListCount - 1

   If Main_Form.ListBox1.ListCount = 0 Then
   
            GoTo Leave_Attachents:
   
   Else
   
           .AddAttachment (Main_Form.ListBox1.List(Index))
         
  
   End If
   
Next

Leave_Attachents:

'Send all
'Set .Configuration = iConf

        .Send
    End With
    XP_ProgressBar1.Value = 60
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing
      'Set iMsg = Nothing
'Set iConf = Nothing
'Set Flds = Nothing
'Set schema = Nothing
    SendEmail = True
    XP_ProgressBar1.Value = 100
    XP_ProgressBar1.Visible = False
    MsgBox "success"
    Exit Function
ErrTrap:
Err.Raise Err.Number, "", "Error from Functions.SendEmail" & Err.Description
    SendEmail = False
End Function
Private Sub Form_Unload(Cancel As Integer)

Dim Response1 As Integer

'If Form canceled with Close button
Cancel = 1

'Confirm if user wants or not
Response1 = MsgBox("Do you really wants to close program?", vbQuestion + vbYesNo, "Confirmation")


If Response1 = vbYes Then

'If user wants then
End

Else

'Else Close message
End If

End Sub

Private Sub jcbutton1_Click()
FormMain.Show

End Sub

Private Sub TextBox10_Change()
If TextBox10.Text = "" Then
MsgBox "Please enter Gmail password.", vbExclamation, "Gmail password"
Else
TextBox5.Visible = True
TextBox4.Visible = True
End If

End Sub

Private Sub TextBox2_Change()
If TextBox2.Text = "" Then
MsgBox "Please enter Receiver's name.", vbExclamation, "Receiver's name"
Else
TextBox3.Visible = True
End If

End Sub

Private Sub TextBox3_Change()
If TextBox3.Text = "" Then
MsgBox "Please enter Sender's name.", vbExclamation, "Sender's name"
Else
TextBox8.Visible = True
End If
End Sub

Private Sub TextBox4_Change()
If TextBox4.Text = "" Then
MsgBox "Please enter Sender's name.", vbExclamation, "Sender's name"
Else
TextBox2.Visible = True

End If
End Sub

Private Sub TextBox8_Change()
If TextBox8.Text = "" Then
MsgBox "Please enter Receiver's name.", vbExclamation, "Receiver's name"
Else
'TextBox8.Text = TextBox8.Text
CommandButton1.Visible = True
End If

End Sub

Private Sub TextBox9_Change()

'Automatically set sender email address
TextBox5.Text = TextBox9.Text
  'Check for a valid gmail address
    If InStr(1, TextBox9.Text, "@gmail.com") < 1 Then
  
  
        MsgBox "Email address is not a valid gmail address. Please enter correct gmail address like sample@gmail.com.", vbOKOnly + vbCritical, App.Title
        TextBox9.SetFocus
        Exit Sub
        Else
        TextBox10.Visible = True
        End If
        
        
End Sub

Private Sub Trim_Functions()

'Setting veriables to trim text of textboxes
Dim Gmail_Address As String
Dim Receiver_email As String
Dim Subject As String
Dim Sender_name As String
Dim Receiver_name As String
Dim CC As String
Dim Bcc As String

'Trimming all
Gmail_Address = Trim$(TextBox9.Text)
Receiver_email = Trim$(TextBox3.Text)
Subject = Trim$(TextBox8.Text)
Sender_name = Trim$(TextBox4.Text)
Receiver_name = Trim$(TextBox2.Text)
CC = Trim$(TextBox6.Text)
Bcc = Trim$(TextBox7.Text)

'Putting trimed text back to textboxes
Main_Form.TextBox9.Text = Gmail_Address
Main_Form.TextBox3.Text = Receiver_email
Main_Form.TextBox8.Text = Subject
Main_Form.TextBox4.Text = Sender_name
Main_Form.TextBox2.Text = Receiver_name
Main_Form.TextBox6.Text = CC
Main_Form.TextBox7.Text = Bcc


End Sub
