VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1260
   ClientLeft      =   5175
   ClientTop       =   5070
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2453
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1253
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Click here to email me"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   773
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project demonstrates how to disable the ALT CTRL DEL buttons.
'You'll also find out how to send email from within your programs.
'You are free to us this code as and how you like.
'Thanks for viewing this example. Don't forget to vote!
'Email: crouchie1998@hotmail.com
'Crouchie1998


        'Function used to disable ALT CTRL DEL
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

        'Function used to send email
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

        'Constant used to disable ALT CTRL DEL
Private Const SPI_SCREENSAVERRUNNING = 97

        'Set the constant caption used in this example
Private Const strCaptions As String = "ALT CTRL DELETE Example"

Dim ret As Integer
Dim bolDisabled As Boolean
Dim strMsg As String

Private Sub Command1_Click()

        'Set the caption of command1 and enables/disables the ALT CTRL DEl buttons
    If Command1.Caption = "Disable" Then
        Command1.Caption = "Re-Enable"
        ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, bolDisabled, 0)
    Else
        Command1.Caption = "Disable"
        ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, bolDisabled, 0)
    End If
    
End Sub

Private Sub Command2_Click()

        'Check to see if ALT CTRL DEL are disabled. If so, re-enable them
    If Not ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, bolDisabled, 0) Then
        ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, bolDisabled, 0)
    End If
    
        'Message to pass to the msgbox below
    strMsg = " ALT CTRL DEL have been re-enabled" & vbCrLf & vbCrLf & _
    vbTab & "  crouchie1998@hotmail.com" & vbCrLf & vbCrLf & vbTab & vbTab & _
    "Don't forget to vote"

        'A message box with the above message
    MsgBox strMsg, vbInformation

        'Closes the program
    End

End Sub

Private Sub Form_Load()

        'Set the form's caption
    Form1.Caption = strCaptions

        'Set the button captions
    Command1.Caption = "Disable"
    Command2.Caption = "Close"

        'Set the label captions
    Label1.Caption = strCaptions
    Label2.Caption = "Email Me"

        'Set the colour of the label to blue to mimic a hyperlink
    Label2.ForeColor = vbBlue

End Sub

Private Sub Form_Unload(Cancel As Integer)
        'Enable the ALT CTRL DEL buttons
    Command2_Click
End Sub

Private Sub SendEmail(strTo As String, strSubject As String)
        'Open the default email client with the strings passed to it.
        'The second string can be ommited
    ret = ShellExecute(0, vbNullString, "mailto:" & strTo & "?subject=" & strSubject, vbNullString, vbNullString, vbnormailfocus)
End Sub

Private Sub Label2_Click()
        'Send the string(s) to the 'SendEmail' sub. You can ommit the subject
    SendEmail "crouchie1998@hotmail.com", "RE: ALT CTRL DEL Example by Crouchie1998. Don't forget to vote!"
End Sub
