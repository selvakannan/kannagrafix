VERSION 5.00
Begin VB.Form Dialog2 
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Desktop_Email_Client.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Label1 
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Dialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Timer1_Timer()

'Attached files counter veriables
Dim i As Integer
Dim Index

'Email Sending veriables
Dim iMsg
Dim iConf
Dim Flds
Dim schema

'Setting Progressbar & Sending Status
Dialog2.XP_ProgressBar1.Value = 20
Dialog2.Label2.Caption = "Connecting to the server . . . . ."

On Error GoTo SendMail_Error:

Set iMsg = CreateObject("CDO.Message")

Set iConf = CreateObject("CDO.Configuration")

Set Flds = iConf.Fields

schema = "http://schemas.microsoft.com/cdo/configuration/"

Flds.Item(schema & "sendusing") = 2

'Server Address (Must be Smtp.gmail.com)
Flds.Item(schema & "smtpserver") = Main_Form.Label10.Caption

'Server Port (465)
Flds.Item(schema & "smtpserverport") = Main_Form.Label11.Caption

Dialog2.XP_ProgressBar1.Value = 30

'Athentication type
Flds.Item(schema & "smtpauthenticate") = 1

'Gmail complete address as Username
Flds.Item(schema & "sendusername") = Main_Form.TextBox9.Text

'Gmail ID password
Flds.Item(schema & "sendpassword") = Main_Form.TextBox10.Text

'Connection timeout
Flds.Item(schema & "smtpConnectionTimeout") = 40

'SSL setting
Flds.Item(schema & "smtpusessl") = 1

Flds.Update

'Show progress of sending
Dialog2.XP_ProgressBar1.Value = 50
Dialog2.Label2.Caption = "Please wait while sending email . . . . ."
'MsgBox
'Setting-up email perameters
MsgBox Main_Form.TextBox2.Text & "<" & Main_Form.TextBox3.Text & ">" & Main_Form.TextBox4.Text & "<" & Main_Form.TextBox5.Text & ">" & Main_Form.TextBox8.Text
With iMsg
   .To = Main_Form.TextBox2.Text & "<" & Main_Form.TextBox3.Text & ">"
   .From = Main_Form.TextBox4.Text & "<" & Main_Form.TextBox5.Text & ">"
   .CC = Main_Form.TextBox6.Text
   .Bcc = Main_Form.TextBox7.Text
   .Subject = Main_Form.TextBox8.Text
   
    'E-mail Text-body
    Dialog2.XP_ProgressBar1.Value = 60
   .TextBody = Main_Form.TextBox1.Text
   

End With



End Sub
Private Sub Clear_Data()

'Clear files list & disable delete button
Main_Form.ListBox1.Clear
Main_Form.CommandButton3.Enabled = False

'Clear Main form's data fields
Main_Form.TextBox1.Text = ""
Main_Form.TextBox2.Text = ""
Main_Form.TextBox3.Text = ""
Main_Form.TextBox6.Text = ""
Main_Form.TextBox7.Text = ""
Main_Form.TextBox8.Text = ""
'TextBox2.Text = ""
'TextBox3.Text = ""
Main_Form.TextBox4.Text = ""
Main_Form.TextBox9.Text = ""
Main_Form.TextBox10.Text = ""
'TextBox8.Text = ""

'Update attached files counter label
Main_Form.Label9.Caption = Main_Form.ListBox1.ListCount & " File (s) attached at this time."

End Sub
