VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNO 
      Caption         =   "No"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1560
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optYES 
      Caption         =   "Yes"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Email Status"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   4455
      Begin VB.Label lblIntMsgs 
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Mail Messages:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "password"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "username"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "pop.server.com"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCheckMail 
      Caption         =   "Check Mail"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   4200
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer sock_timer 
      Interval        =   125
      Left            =   3960
      Top             =   3120
   End
   Begin VB.Label Label5 
      Caption         =   "Mark Emails for Deletion off POP Host?"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "POP PASSWORD:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "POP USERNAME:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "POP MAIL SERVER:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    POP3_TOP
    POP3_RETR
    POP3_DELE
    POP3_QUIT
End Enum
Private Type POP_MAIL_TYPE
 intMessages As Integer
 intCurrentMessage As Integer
 intCurrentHeader As Integer
 intCurrentDelete As Integer
 Buffer As String
 hdrBuffer As String
 Headers() As String
 Message() As String
End Type
Private m_State As POP3States
Private pop As POP_MAIL_TYPE

Private Sub cmdCheckMail_Click()
 If Len(txtServer) = 0 Then
  MsgBox "You must define a Pop server", vbCritical
  txtServer.SetFocus
  Exit Sub
 ElseIf Len(txtUser) = 0 Then
  MsgBox "You must give a username", vbCritical
  txtUser.SetFocus
  Exit Sub
 ElseIf Len(txtPass) = 0 Then
  MsgBox "You must give a password", vbCritical
  txtPass.SetFocus
  Exit Sub
 End If
 
 ' Set Session State
 m_State = POP3_Connect
 ' Connect to Server
 sock.Close
 sock.Connect txtServer, 110
 
txtServer.Enabled = False
txtUser.Enabled = False
txtPass.Enabled = False
End Sub

Private Sub cmdExit_Click()
sock.Close
End
End Sub



Private Sub lblIntMsgs_Click()
MsgBox UBound(pop.Message)
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
sock.GetData Data

  If Left(Data, 1) = "+" Or m_State = POP3_RETR Or m_State = POP3_TOP Then
    Select Case m_State
    
      Case POP3_Connect
      'Reset Number of messages
       pop.intMessages = 0
      'Change Session State
      m_State = POP3_USER
      ' send the User Name
      sock.SendData "USER " & txtUser & vbCrLf
      
      Case POP3_USER
      'Change Session State
      m_State = POP3_PASS
      'Send PassWord
      sock.SendData "PASS " & txtPass & vbCrLf
      
      Case POP3_PASS
      'Change Session State
      m_State = POP3_STAT
      'Send Stat Command
      sock.SendData "STAT" & vbCrLf
      
      Case POP3_STAT
       pop.intMessages = CInt(Mid$(Data, 5, _
                         InStr(5, Data, " ") - 5))
        'Show status on form
        lblIntMsgs = pop.intMessages
        If pop.intMessages > 0 Then
          ' You have Mail
          ' Set message array
          ReDim pop.Message(1 To pop.intMessages)
          '' Now get all the headers
          ReDim pop.Headers(1 To pop.intMessages)
          ' reset state
          m_State = POP3_TOP
        'increment header counter
        pop.intCurrentHeader = pop.intCurrentHeader + 1
        sock.SendData "TOP " & pop.intCurrentHeader & " 0" & vbCrLf
        Else
         'Mailbox is empty
         ' reset state
         m_State = POP3_QUIT
         'send quit command
         sock.SendData "QUIT" & vbCrLf
        End If
        
      Case POP3_TOP

       pop.hdrBuffer = pop.hdrBuffer & Data
       If InStr(1, pop.hdrBuffer, vbCrLf & "." & vbCrLf) Then
       'Delete First Line in response
       pop.hdrBuffer = Mid$(pop.hdrBuffer, InStr(1, pop.hdrBuffer, vbCrLf) + 2)
       'Delete Last Line (the ".")
       pop.hdrBuffer = Left$(pop.hdrBuffer, Len(pop.hdrBuffer) - 3)
       ' add message to array
       pop.Headers(pop.intCurrentHeader) = pop.hdrBuffer
       'Clear buffer
       pop.hdrBuffer = ""
        'increment header counter
         If pop.intCurrentHeader = pop.intMessages Then
         'Change Session State
          m_State = POP3_RETR
          'increment number of messages by one
          pop.intCurrentMessage = pop.intCurrentMessage + 1
          ' Send request for 1st message
          sock.SendData "RETR 1" & vbCrLf
       Else
        pop.intCurrentHeader = pop.intCurrentHeader + 1
        sock.SendData "TOP " & pop.intCurrentHeader & " 0" & vbCrLf
       End If
        End If

       
      Case POP3_RETR
       pop.Buffer = pop.Buffer & Data
       If InStr(1, pop.Buffer, vbCrLf & "." & vbCrLf) Then
       'Message Done
       'Delete First Line in response
       pop.Buffer = Mid$(pop.Buffer, InStr(1, pop.Buffer, vbCrLf) + 2)
       'Delete Last Line (the ".")
       pop.Buffer = Left$(pop.Buffer, Len(pop.Buffer) - 3)
       ' add message to array
       pop.Message(pop.intCurrentMessage) = pop.Buffer
       'Clear buffer
       pop.Buffer = ""
         If pop.intCurrentMessage = pop.intMessages Then
          ' if equal, there are no more messages
          ' so lets mark them for deletion if
          ' user wants
          'Reset State
          If optNO = True Then
          ' set session for quit
          m_State = POP3_QUIT
          ' Send QUIT command
          sock.SendData "QUIT" & vbCrLf
          Else
          ' set session for delete function
           m_State = POP3_DELE
           'Send the NOOP command just for a
           ' check to make sure we still have good connection
           sock.SendData "NOOP" & vbCrLf
          End If
         Else
          ' There are more messages
          ' increment counter
          pop.intCurrentMessage = pop.intCurrentMessage + 1
          'Reset Session State
          m_State = POP3_RETR
          ' Send Request for next message
          sock.SendData "RETR " & _
          CStr(pop.intCurrentMessage) & vbCrLf
         End If
       End If
       
     Case POP3_DELE
       ' Check to see where we are
       If pop.intCurrentDelete = pop.intMessages Then
         'if true then we are done
         'reset state
         m_State = POP3_QUIT
         'Send Quit command
         sock.SendData "QUIT" & vbCrLf
       Else
        'Increment Delete counter
        pop.intCurrentDelete = pop.intCurrentDelete + 1
        ' We still have messages to Delete
        sock.SendData "DELE " & pop.intCurrentDelete & vbCrLf
       End If
       
     Case POP3_QUIT
      'Finished with whatever we received
      sock.Close
      'Reset Some Stuff
      pop.Buffer = ""
      pop.intCurrentMessage = 0
      pop.intMessages = 0
      pop.intCurrentHeader = 0
      pop.intCurrentDelete = 0
      'Now We Need to parse our email messages
      ' if there are any
      On Error GoTo Uerr:
      If UBound(pop.Message) > 0 Then
       ParseEmail pop.Message, pop.Headers
      End If
Uerr:
      'Clear from Memory
      ReDim pop.Message(0)
      ReDim pop.Headers(0)
      
      txtServer.Enabled = True
      txtUser.Enabled = True
      txtPass.Enabled = True
      
    End Select
  Else
   '' Error so close
   sock.Close
   MsgBox "POP ERROR" & vbCrLf & Data, vbExclamation
  End If
End Sub

Private Sub sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print Description
sock.Close
End Sub

Private Sub sock_timer_Timer()
lblStatus = SockStatus(sock)
End Sub

Private Sub txtPass_GotFocus()
txtPass.SelStart = 0
txtPass.SelLength = Len(txtPass)
End Sub
Private Sub txtUser_GotFocus()
txtUser.SelStart = 0
txtUser.SelLength = Len(txtUser)
End Sub
Private Sub txtServer_Change()
If sock.State <> sckConnected Then
lblIntMsgs = 0
End If
End Sub

Private Sub txtServer_GotFocus()
txtServer.SelStart = 0
txtServer.SelLength = Len(txtServer)
End Sub
