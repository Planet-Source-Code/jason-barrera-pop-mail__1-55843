Attribute VB_Name = "POPMAIL"
Public Type Msgs
 cnt As Integer
 pos As Integer
 html() As Boolean
 Boundary() As String
 Headers() As String
 mTo() As String
 From() As String
 Subject() As String
 Date() As String
 Body() As String
 Email() As String
 attFileName() As String
 attFile() As String
End Type
Public e_msgs As Msgs
Dim AttCount As Long
Public Function ParseEmail(ByVal Message As Variant, ByVal Headers As Variant)
''Do Defaults''
e_msgs.cnt = UBound(Message)
ReDim e_msgs.Email(LBound(Message) To UBound(Message))
ReDim e_msgs.Headers(LBound(Message) To UBound(Message))
ReDim e_msgs.Body(LBound(Message) To UBound(Message))
ReDim e_msgs.Boundary(LBound(Message) To UBound(Message))
ReDim e_msgs.html(LBound(Message) To UBound(Message))
ReDim e_msgs.mTo(LBound(Message) To UBound(Message))
ReDim e_msgs.From(LBound(Message) To UBound(Message))
ReDim e_msgs.Subject(LBound(Message) To UBound(Message))
ReDim e_msgs.Date(LBound(Message) To UBound(Message))
ReDim e_msgs.attFile(LBound(Message) To UBound(Message))
ReDim e_msgs.attFileName(LBound(Message) To UBound(Message))

'' Dim some Specific Vars
Dim Boundary As String
Dim i As Integer
Dim posA As Long
Dim bPos As Long
Dim bPos2 As Long
'' Execute
For i = LBound(Message) To UBound(Message)
 '' parse headers for to, from, subject etc..
 Dim sp() As String
 sp = Split(Headers(i), vbCrLf)
 Dim s As Variant
 For Each s In sp
  If Left(s, 4) = "To: " Then
   If Len(e_msgs.mTo(i)) = 0 Then
   e_msgs.mTo(i) = Mid(s, 5)
   End If
  End If
  If Left(s, 6) = "From: " Then
   If Len(e_msgs.From(i)) = 0 Then
   e_msgs.From(i) = Mid(s, 7)
   End If
  End If
  If Left(s, 9) = "Subject: " Then
   If Len(e_msgs.Subject(i)) = 0 Then
   e_msgs.Subject(i) = Mid(s, 10)
   End If
  End If
  If Left(s, 6) = "Date: " Then
   If Len(e_msgs.Date(i)) = 0 Then
   e_msgs.Date(i) = Mid(s, 7)
   End If
  End If
 Next
 
posA = InStr(1, Message(i), vbCrLf & vbCrLf)
If posA > 0 Then
'' GEt Boundary
bPos = InStr(1, Headers(i), "boundary=")
 If bPos > 0 Then
 bPos = InStr(bPos, Headers(i), """") + 1
 bPos2 = InStr(bPos, Headers(i), """")
 Boundary = Mid(Headers(i), bPos, bPos2 - bPos)
 e_msgs.Boundary(i) = Boundary
 Else
 e_msgs.Boundary(i) = ""
End If

'' Strip Boundary from bottom of body
 bPos = InStrRev(Message(i), "--" & Boundary) - 1
If bPos > 0 Then
 e_msgs.Body(i) = Replace(Left$(Message(i), bPos), Headers(i), "")
 bPos = Len(e_msgs.Body(i)) - 3
  If InStr(bPos, e_msgs.Body(i), vbCrLf) Then
   bPos = InStr(bPos, e_msgs.Body(i), vbCrLf)
   e_msgs.Body(i) = Left$(e_msgs.Body(i), bPos - 1)
  End If
 Else
 e_msgs.Body(i) = Replace(Message(i), Headers(i), "") 'Right$(Message(i), (Len(Message(i)) - posA) - 3)
End If

''parse for attachments
'How many?
AttCount = StrCount(e_msgs.Body(i), "Content-Disposition: attachment;", False)
bPos = InStr(1, e_msgs.Body(i), "Content-Disposition: attachment;", vbTextCompare)
If bPos Then
 bPos = bPos + 32
 bPos = InStr(bPos, e_msgs.Body(i), "filename=" & """", vbTextCompare)
 If bPos Then
  bPos = bPos + 10
  bPos2 = InStr(bPos, e_msgs.Body(i), """")
  If bPos2 Then
   e_msgs.attFileName(i) = Mid(e_msgs.Body(i), bPos, bPos2 - bPos)
   bPos = InStr(bPos2, e_msgs.Body(i), vbCrLf & vbCrLf)
   If bPos Then
    bPos = bPos + 4
    bPos2 = InStr(bPos, e_msgs.Body(i), "==")
    If bPos2 Then
     e_msgs.attFile(i) = Base64Decode(Mid(e_msgs.Body(i), bPos, (bPos2 + 2) - bPos))
     SaveMail e_msgs.attFile(i), App.Path & "\SaveMail\" & e_msgs.attFileName(i)
    End If
   End If
  End If
 End If
End If


'' Parse for content type
e_msgs.html(i) = False
bPos = InStr(1, e_msgs.Body(i), "<html>", vbTextCompare)
If bPos > 0 Then
 bPos2 = InStr(bPos, e_msgs.Body(i), "</html>", vbTextCompare)
 If bPos2 > 0 Then
  e_msgs.html(i) = True
  e_msgs.Body(i) = Mid(e_msgs.Body(i), bPos, bPos2 + 7)
 Else
 e_msgs.Body(i) = Mid(e_msgs.Body(i), bPos)
 End If
End If

 e_msgs.Headers(i) = Headers(i) 'Left$(Message(i), posA - 1)

 
 e_msgs.Email(i) = Message(i)
Else
 e_msgs.Headers(i) = "INVALID FORMAT"
 e_msgs.Body(i) = "INVALID FORMAT"
 e_msgs.Email(i) = Message(i)
End If

 Message(i) = ""

 'Save to Disk
 SaveMail e_msgs.Email(i), App.Path & "\SaveMail\eMsg_" & i & ".txt"
Next i
ReDim Message(0)
Form2.Show
End Function
Public Function SaveMail(ByVal Mail As String, ByVal Path As String)
Dim Handle As Integer
Handle = FreeFile
Open Path For Binary As #Handle
 Put #Handle, , Mail
Close #Handle
End Function
