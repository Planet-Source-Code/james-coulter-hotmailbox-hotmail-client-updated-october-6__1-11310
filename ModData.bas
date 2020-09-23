Attribute VB_Name = "Module"
''''''''''''''''''''''''''''' Hotmail Check Message ''''''''''''''''''''''''''''
'                                                                              '
'    This code uses the http/1.1 protocol to connect to the hotmail server     '
'    and retrieve the mail box (note: when i use the term mailbox 'data'       '
'    I am actually referring to the SOURCE CODE of the mailbox, which of       '
'    course is sent in html format). This program does not use any special     '
'    mail features, nor does it implement POP mail, it simply uses http        '
'    commands to get the mailbox. Because it is so confusing, I tried the      '
'    best i could to comment anywhere that there may be confusion, but         '
'    if you are not familiar with socket programming or the http protocol,     '
'    you will most likely have a difficult time understanding it.              '
'    And although the only piece of data you see as a result of this program   '
'    is how many new messages you have, once you understand how the program    '
'    works, retrieving any other information about your hotmail account is     '
'    a piece of cake. If you have any questions or comments, you can contact   '
'    me at:  nmjblue@hotmail.com                                               '
'                                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public NextPage As String, CurrentPage As String, FolderURL As String, msgURL As String
Public msgIDX As Long
Public ShowDlgs As Boolean, Interval As Integer, lastmsg As Long
Public StrLogin As String, StrPass As String ' holds login and password
Public NewHost As String, NewUrl As String ' new server and url after redirection (see below)
Public BatchNumber As Integer ' holds the current batch number we need to send
Public Cookies(6) As String ' stores cookies received, required for receiving mailbox (contains encrypted information read by server)
Public CurrentCookie As Integer ' stores current cookie number, as there are numerous different ones
Public MailData As String ' once we begin to receive data about mailbox, this is the string that stores it so we can retrieve the information
Public ReadBox As Boolean, BoxBatch As Integer ' boolean for whether or not we are receiving the mailbox data, and batch number of the data we are receiving
Public loggedin As Boolean
Public GotMail As Boolean, composeString As String
Public composeurl As String, newmessages As String
Public newmail As Long, msgfolder As String


Public msghdrid As String, sendURL As String

' Socket Values
Public Const AF_INET = 2
Public Const SOCK_STREAM = 1
Public Const IPPROTO_IP = 0
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_DISCONNECT = 7

Public Type HotmailMsg
    subject As String
    sender As String
    email As String
    indate As String
    newmail As Boolean
    msgURL As String
    attach As Boolean
    attachURL As String
    size As Long
    index As Long
    cached As Boolean
    ' Indicates whether the message is still online
    isonline As Boolean
    ' Indicates whether the message is new as far as the program is concerned
    isnew As Boolean
End Type

Public Type HotmailFolder
    fname As String
    id As String
    url As String
    size As Long
    msgs As Long
    newmsgs As Long
End Type

Public Type HotmailAddress
    fullname As String * 128
    email As String * 128
End Type

Public Type HotmailAccount
    username As String * 64
    loginname As String * 64
    ' This will have to be encrypted someday
    password As String * 12
End Type

Public Type upVersion
    major As Integer
    minor As Integer
    rev As Integer
End Type

'Public NewMessages As Integer

Public Folders() As HotmailFolder
Public FolderCount As Long

Public Messages() As HotmailMsg
Public MsgCount As Long

Public Addresses() As HotmailAddress
Public addcount As Long

Public Accounts() As HotmailAccount
Public AccCount As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpSound As String, ByVal flag As Long) As Long

Public Function AddMessage(ByVal subject As String, ByVal sender As String, ByVal email As String, Optional ByVal indate As String, Optional newmsg As Boolean, Optional ByVal url As String) As Long
    Dim i As Long
    Dim lastnum As Long
    ' Make sure the message is not already loaded
    For i = 0 To MsgCount - 1
        If InStr(1, Messages(i).subject, Left(subject, Len(subject) - 3)) = 1 And Messages(i).indate = indate And Messages(i).email = email Then
            Messages(i).isonline = True
            Messages(i).newmail = newmsg
            If Not Messages(i).cached Then
                Messages(i).isnew = True
            End If
            AddMessage = i
            Exit Function
        End If
        
        If Messages(i).index > lastnum Then
            lastnum = Messages(i).index
        End If
    Next
        
    newmail = newmail + 1
        
    ReDim Preserve Messages(MsgCount) As HotmailMsg
    Messages(MsgCount).isonline = True
    With Messages(MsgCount)
        .subject = subject
        .sender = sender
        .email = email
        .indate = indate
        .newmail = newmsg
        .isnew = True
        .msgURL = url
        .index = lastnum + 1
    End With
    AddMessage = MsgCount
    MsgCount = MsgCount + 1
End Function

Public Function MakeSendString(ByVal tomail As String, ByVal subject As String, ByVal body As String, Optional ByVal cc As String, Optional ByVal bcc As String, Optional ByVal signature As String)
    Dim content As String
    Dim strdata As String ' for temporary storage of data to send
    Dim feed As String
    feed = (Chr(13) & Chr(10))
    
    ' We have to first take out all the spaces and replace them with plus signs
    Dim tt1 As Long
    Dim part1 As String, part2 As String
    Dim last As Long
    
    body = body & vbCrLf & signature
    
    tt1 = InStr(1, body, " ")
    Do Until tt1 = 0
        part1 = Mid(body, 1, tt1 - 1)
        part2 = Mid(body, tt1 + 1, Len(body) - tt1)
            
        body = part1 & "+" & part2
        
        last = tt1 + 1
        tt1 = InStr(1, body, " ")
    Loop
    
    If frmhotmail.chkSaveSent.Value = 1 Then
        content$ = "login=" & StrLogin & "&wcid=&msg=&start=&len=&attfile=&type=&src=&subaction=&wysiwyg=&ref=&sigflag=y&newmail=new&msghdrid=" & msghdrid & "&col_name=Name&col_size=Size&col_type=Type&col_mod=Modified&col_path=Path&dlog_choosefile=Choose+File&dlog_progress=Attachment+Upload+Progress&dlog_delete=Attachments+on+the+server+can+not+be+removed.&dlog_zerok=Cannot+send+a+zero+length+attachment.&dlog_sizeexceeded=The+total+size+of+attachments+cannot+exceed+1000k.&dlog_filenotfound1=The+file+&dlog_filenotfound2=+could+not+be+found.+Continue%3F&dlog_badserver=The+target+server+is+not+a+valid+Hotmail+server.&to=" & tomail & "&subject=" & subject & "&cc=" & cc & "&bcc=" & bcc & "&outgoing=on&Send.x=Send&body=" & body & "&TMP_outgoing=on"
    Else
        content$ = "login=" & StrLogin & "&wcid=&msg=&start=&len=&attfile=&type=&src=&subaction=&wysiwyg=&ref=&sigflag=y&newmail=new&msghdrid=" & msghdrid & "&col_name=Name&col_size=Size&col_type=Type&col_mod=Modified&col_path=Path&dlog_choosefile=Choose+File&dlog_progress=Attachment+Upload+Progress&dlog_delete=Attachments+on+the+server+can+not+be+removed.&dlog_zerok=Cannot+send+a+zero+length+attachment.&dlog_sizeexceeded=The+total+size+of+attachments+cannot+exceed+1000k.&dlog_filenotfound1=The+file+&dlog_filenotfound2=+could+not+be+found.+Continue%3F&dlog_badserver=The+target+server+is+not+a+valid+Hotmail+server.&to=" & tomail & "&subject=" & subject & "&cc=" & cc & "&bcc=" & bcc & "&body=" & body & "&Send.x=Send"
    End If
    strdata = "POST " & sendURL & " HTTP/1.1" & feed & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*" & feed
    strdata = strdata & "Accept -Language: en -us" & feed
    strdata = strdata & "Referer: http://" & NewHost & "/cgi-bin/compose?a=b" & feed
    strdata = strdata & "Accept -Encoding: gzip , deflate" & feed & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
    strdata = strdata & "Content-Type: application/x-www-form-urlencoded" & feed
    strdata = strdata & "Host: " & NewHost & feed
    strdata = strdata & "Content-Length: " & Len(content$) & feed & "Connection: Keep -Alive" & feed
    strdata = strdata & "Cookie: HMP1=1; " & Cookies(4) & "; " & Cookies(1) & "; " & Cookies(2) & feed & feed
    strdata = strdata & content$ & feed & feed
    MakeSendString = strdata
End Function

Public Function MakeString(Connection As Long, Optional ByVal url As String) As String
    Dim strdata As String ' for temporary storage of data to send
    Dim feed As String
    feed = (Chr(13) & Chr(10)) ' carriage return & linefeed
    
    Select Case Connection
        Case 0 'first batch of data sent, contains login information
            Dim content As String
            content$ = "login=" & StrLogin$ & "&domain=hotmail.com&passwd=" & StrPass$ & "&enter=Sign+in&sec=no&curmbox=ACTIVE&js=yes&_lang=&beta=&ishotmail=1&id=2&ct=963865176"
            strdata = "POST /cgi-bin/dologin HTTP/1.1" & feed & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*" & feed
            strdata = strdata & "Accept -Language: en -us" & feed & "Content-Type: application/x-www-form-urlencoded" & feed
            strdata = strdata & "Accept -Encoding: gzip , deflate" & feed & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
            strdata = strdata & "Host: lc5.law5.hotmail.passport.com" & feed
            strdata = strdata & "Content-Length: " & Len(content$) & feed & "Connection: Keep -Alive" & feed & feed
            strdata = strdata & content$ & feed & feed
            MakeString = strdata
        Case 1 'we get relocated to a new hotmail server (NewHost) containing the mailbox. here we request a new page, because contained in the url of the page (NewUrl) is our encrypted login and password
            strdata = "GET /" & NewUrl$ & " HTTP/1.1" & feed
            strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
            strdata = strdata & "Host: " & NewHost$ & feed
            strdata = strdata & "Cookie: MC1=V=2&GUID=B8E9C518070C49B18A9884F543033C33; mh=ENCA; MSPDom=; MSPAuth=; MSPProf=; MSPVis=; LO=; HMSC0899=; HMP1=1; HMSC0899="
            strdata = strdata & feed & feed
             '& feed
            MakeString = strdata
        Case 2 'finally, we request the mailbox on the new server, by sending the cookies we received with all the encrypted information needed
            strdata = "GET " & NewUrl$ & " HTTP/1.1" & feed
            strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
            strdata = strdata & "Host: " & NewHost$ & feed
            strdata = strdata & "Connection: Keep-Alive" & feed
            strdata = strdata & "Cookie: HMP1=1; " & Cookies(4) & "; MSPDom=; " & Cookies(1) & "; " & Cookies(2) & "; MSPVis=1; LO=;" & feed & feed
            MakeString = strdata
        Case 3
            ' Get Next Page of Messages
            strdata = "GET " & url & " HTTP/1.1" & feed & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*" & feed
            'If LastPage <> "" Then
            strdata = strdata & "Referer: http://" & NewHost & CurrentPage & feed
            'End If
            strdata = strdata & "Accept-Language: en-us" & feed '& "Content-Type: application/x-www-form-urlencoded" & feed
            strdata = strdata & "Accept-Encoding: gzip, deflate" & feed
            strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
            strdata = strdata & "Host: " & NewHost$ & feed
            strdata = strdata & "Connection: Keep-Alive" & feed
            strdata = strdata & "Cookie: HMP1=1; " & Cookies(4) & "; " & Cookies(1) & "; " & Cookies(2) & feed & feed
            MakeString = strdata
        Case 4
            ' Get Folder List
            strdata = "GET /cgi-bin/" & FolderURL & " HTTP/1.1" & feed & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*" & feed
            strdata = strdata & "Referer: http://" & NewHost & NextPage & feed
            strdata = strdata & "Accept-Language: en-us" & feed '& "Content-Type: application/x-www-form-urlencoded" & feed
            strdata = strdata & "Accept-Encoding: gzip, deflate" & feed
            strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
            strdata = strdata & "Host: " & NewHost$ & feed
            strdata = strdata & "Connection: Keep-Alive" & feed
            strdata = strdata & "Cookie: HMP1=1; " & Cookies(4) & "; " & Cookies(1) & "; " & Cookies(2) & feed & feed
            MakeString = strdata
        Case 5
            ' Get Message Body
            strdata = "GET /cgi-bin/" & url & " HTTP/1.1" & feed
            strdata = strdata & "Referer: http://" & NewHost & NextPage & feed
            strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
            strdata = strdata & "Host: " & NewHost$ & feed
            strdata = strdata & "Connection: Keep-Alive" & feed
            strdata = strdata & "Cookie: HMP1=1; " & Cookies(4) & "; MSPDom=; " & Cookies(1) & "; " & Cookies(2) & "; MSPVis=1; LO=;" & feed & feed
            MakeString = strdata
        Case 6
            ' Delete Message
            content = "tobox=&js=&_HMaction=delete&foo=inbox&page=&" & url & "=on&nullbox="
            strdata = strdata & "POST /cgi-bin/HoTMaiL HTTP/1.1" & feed & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*" & feed
            strdata = strdata & "Referer: http://" & NewHost & NextPage & feed
            strdata = strdata & "Accept-Language: en-us" & feed & "Content-Type: application/x-www-form-urlencoded" & feed
            strdata = strdata & "Accept-Encoding: gzip, deflate" & feed
            strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
            strdata = strdata & "Host: " & NewHost$ & feed
            strdata = strdata & "Content-Length: " & Len(content$) & feed & "Connection: Keep-Alive" & feed
            strdata = strdata & "Cookie: HMP1=1; " & Cookies(4) & "; MSPDom=; " & Cookies(1) & "; " & Cookies(2) & "; MSPVis=1; LO=;" & feed & feed
            strdata = strdata & content & feed & feed
            MakeString = strdata
    End Select
End Function

Public Sub PlayWave(sFileName As String)
    On Error GoTo Play_Err
    
    Dim iReturn As Integer
    
    If frmhotmail.chkNoSounds.Value = 0 Then
        'Make sure something was passed to the Play Function
        If sFileName > "" Then
            'Make sure a WAV filename was passed
            If UCase$(Right$(sFileName, 3)) = "WAV" Then
                'Make sure the file exists
                If Dir(sFileName) > "" Then
                    iReturn = sndPlaySound(sFileName, 0)
                End If
            End If
        End If
    End If
    
    'Wav file play successful
    Exit Sub

Play_Err:
    'If there was an error then exit without playing
    Exit Sub
End Sub

Public Function FindNextMessage(ByVal start As Long, ByVal text As String) As String
    Dim pos1 As Long, pos2 As Long, pos3 As Long
    Dim epos1 As Long
    Dim temp As String
    Dim sender As String, email As String
    Dim newmail As Boolean
    Dim msgURL As String
    Dim mDate As String
    Dim newmsg As Long, size As String
    
    ' 1. First, we find the next HREF tag that starts with "msg=MSG".  This denotes a Message
    pos1 = InStr(start, text, "msg=MSG")
    If pos1 <> 0 Then
        
        ' GET THE SENDER OF THE MESSAGE
        ' The sender's name is embedded in an HREF tag that is just after the msg=MSG...
        ' So, we should first look for the next ">" and then find the next "<"
        ' and get the part in between.
        
                ' Find the next ">"
                pos2 = InStr(pos1, text, ">")
                If pos2 <> 0 Then
                    ' Find the next "<"
                    pos3 = InStr(pos2, text, "<")
                    If pos3 <> 0 Then
                        ' Extract the sender's name (it's sometimes truncated)
                        sender = RemoveNBSPs(Mid(text, pos2 + 1, (pos3 - pos2) - 1), True)
                    End If
                End If
                
        ' Let's say we want their e-mail address
        ' that information is embedded in a name tag just before the sender's name
        ' we have to be careful about which one we grab though.  It has to be following
        ' a name tag that starts with "MSG"
        epos1 = InStr(start, text, "name=" & Chr(34) & "MSG")
        If epos1 <> 0 Then
            ' Now it should be the next name tag
            pos2 = InStr(epos1 + Len("name=" & Chr(34) & "MSG"), text, "name=")
            If pos2 <> 0 Then
                ' and we'll find the ">"
                pos3 = InStr(pos2, text, ">")
                If pos3 <> 0 Then
                    pos2 = pos2 + Len("name=")
                    email = Mid(text, pos2 + 1, (pos3 - pos2) - 1)
                    ' And take the quotes off the ends
                    email = Left(email, Len(email) - 1)
                    'email = Mid(email, 1, Len(email) - 1)
                End If
            End If
        End If
        
        ' We can also check if the message is new by looking for "newmail.gif"
        ' just so we don't mark the wrong message as new, we have to make sure it
        ' is positioned before the email address we just got
        pos2 = InStr(start, text, "alt='New'")
        If pos2 <> 0 Then
            If pos2 < pos3 And pos2 > epos1 Then
                newmail = True
                'lastmsg = pos2 + 1
            End If
        End If
        
        ' We also want the message URL so we can download the message
        pos2 = InStr(start, text, "getmsg?")
        If pos2 <> 0 Then
            ' find the ending ">"
            pos3 = InStr(pos2, text, ">")
            If pos3 <> 0 Then
                ' extract the url
                msgURL = Mid(text, pos2, (pos3 - 1) - pos2)
            End If
        End If
        
        ' 2. Now, we look for the first <TD> after that
        pos2 = InStr(pos1, text, "<td>")
        If pos2 <> 0 Then
            ' 3. Now we get the </TD> that goes with it
            pos3 = InStr(pos2, text, "</td>")
            If pos3 <> 0 Then
                ' 4. Extract the stuff in between
                pos2 = pos2 + Len("<td>")
                temp = Mid(text, pos2 + 1, pos3 - pos2)
                ' 5. Take 6 characters off of each end.  These are "&nbsp;" tags
                temp = Mid(temp, 6, Len(temp) - 7)
                temp = Left(temp, Len(temp) - 5)
                
                lastmsg = pos3 + 1
                FindNextMessage = temp & " from " & sender & " (" & email & ")"
            End If
        End If
        
        ' The date for the message is embedded in the next <TD>
        pos1 = InStr(pos3 + 1, text, "<td>")
        If pos1 <> 0 Then
            pos2 = InStr(pos1 + 1, text, "</td>")
            If pos2 <> 0 Then
                pos1 = pos1 + Len("<td>")
                mDate = Mid(text, pos1, pos2 - pos1)
                ' We'll have to clean out the &nbsps
                mDate = RemoveNBSPs(mDate)
            End If
        End If
        
        ' The size of the message would be useful when downloading them
        pos1 = InStr(pos2 + Len("</td>") + 1, text, ">")
        If pos1 <> 0 Then
            pos3 = InStr(pos1, text, "</td>")
            If pos3 <> 0 Then
                size = RemoveNBSPs(Mid(text, pos1 + 1, pos3 - pos1), True)
                size = Left(size, Len(size) - 2)
            End If
        End If
                       
        newmsg = AddMessage(temp, sender, email, mDate, newmail, msgURL)
        Messages(newmsg).size = CLng(size)
    End If
End Function

Private Function RemoveNBSPs(ByVal text As String, Optional sp As Boolean) As String
    Dim p1 As Long, p2 As Long
    Dim part1 As String, part2 As String
    
    p1 = InStr(1, text, "&nbsp;")
    Do Until p1 = 0
        part1 = Mid(text, 1, p1 - 1)
        part2 = Mid(text, p1 + Len("&nbsp;"), Len(text) - p1)
        
        If sp Then
            text = part1 & "" & part2
        Else
            text = part1 & " " & part2
        End If
        'last = tt1 + 1
        p1 = InStr(1, text, "&nbsp;")
    Loop
    RemoveNBSPs = Trim(text)
End Function

Public Function GetFolderURL(ByVal text As String) As String
    Dim p1 As Long, p2 As Long
    Dim temp As String
    
    ' Look for folders?a=
    p1 = InStr(1, text, "folders?")
    If p1 <> 0 Then
        ' find the next "
        p2 = InStr(p1, text, Chr(34))
        If p2 <> 0 Then
            temp = Mid(text, p1, p2 - p1)
            GetFolderURL = temp
        End If
    End If
End Function

Public Function ResizeMessageArray()
    ' This function will remove any messages that have been deleted
    ' and update the message count so we don't have wierd things going on
    Dim tempar() As HotmailMsg
    Dim count As Long, i As Long
    
    For i = 0 To MsgCount - 1
        If Messages(i).msgURL <> "" And Messages(i).isonline = True Then
            ReDim Preserve tempar(count) As HotmailMsg
            tempar(count) = Messages(i)
            'tempar(count).index = count
            count = count + 1
        End If
    Next
    
    MsgCount = count
    ReDim Messages(count - 1) As HotmailMsg
    Messages = tempar
End Function

Public Function ProcessFolders(ByVal text As String)
    Dim id1 As Long, id2 As Long
    Dim nm1 As Long
    Dim temp As String
    Dim fname As String
    Dim msgs As Long
    Dim newmsgs As Long
    Dim size As Long
    
    id1 = InStr(1, text, "<tbody>")
    If id1 <> 0 Then
    
        ' Find the first ID tag
        id1 = InStr(id1 + 1, text, "/cgi-bin/HoTMaiL?")
        Do Until id1 = 0
        'If id1 <> 0 Then
            ' Now, find the corresponding "
            id2 = InStr(id1, text, Chr(34))
            If id2 <> 0 Then
                ' Extract the ID
                temp = Mid(text, id1, id2 - id1)
                
                ' Now, we might want the name of the folder as well
                ' to get this, we have to look for a carriage return after the ">"
                nm1 = InStr(id2, text, "<")
                If nm1 <> 0 Then
                    fname = Mid(text, id2 + 2, (nm1 - id2) - 2)
                    If fname <> "" Then
                        ' Now we can get the number of messages
                        ' by looking for the following "center>"
                        id1 = InStr(nm1, text, "center>")
                        If id1 <> 0 Then
                            id2 = InStr(id1 + 1, text, "</")
                            If id2 <> 0 Then
                                id1 = id1 + Len("center>")
                                msgs = CInt(Mid(text, id1, id2 - id1))
                            End If
                        End If
                        ' Make sure we have a starting point if we didn't find
                        ' the number of messages
                        If id2 = 0 Then id2 = nm1
                        ' Do the same for the number of new messages
                        id1 = InStr(id2 + 1, text, "center>")
                        If id1 <> 0 Then
                            id2 = InStr(id1 + 1, text, "</")
                            If id2 <> 0 Then
                                id1 = id1 + Len("center>")
                                If Mid(text, id1, 3) = "<b>" Then
                                    id1 = id1 + 3
                                End If
                                newmsgs = CInt(Mid(text, id1, id2 - id1))
                            End If
                        End If
                        ' And now we do the same for the size of the folder
                        If id2 = 0 Then id2 = nm1
                        id1 = InStr(id2 + 1, text, "right>")
                        If id1 <> 0 Then
                            id2 = InStr(id1 + 1, text, "</")
                            If id2 <> 0 Then
                                id1 = id1 + Len("right>")
                                id2 = id2 - 1
                                size = CLng(Mid(text, id1, id2 - id1))
                            End If
                        End If
                        
                        ReDim Preserve Folders(FolderCount) As HotmailFolder
                        With Folders(FolderCount)
                            .fname = fname
                            .id = temp
                            .url = temp
                            .size = size
                            .newmsgs = newmsgs
                            .msgs = msgs
                        End With
                        FolderCount = FolderCount + 1
                    End If
                End If
            End If
            id1 = InStr(id1 + 1, text, "/cgi-bin/HoTMaiL?")
        'End If
        Loop
    End If
End Function

Public Function IsNextPage(ByVal text As String) As String
    Dim ref1 As Long, ref2 As Long
    Dim t1 As Long
    Dim temp As String
    
    ' The NEXT PAGE text is embedded in an HREF tag that looks like this
    ' <a href="/cgi-bin/HoTMaiL?a=b&page=2">Next Page</a>
    ' So, we will first find the "HoTMaiL?a=b&page=" section
    
    ref1 = InStr(1, text, "/cgi-bin/HoTMaiL?")
    Do Until ref1 = 0
'    If ref1 <> 0 Then
        ' Now we check for the "<"
        ref2 = InStr(ref1, text, "<")
        If ref2 <> 0 Then
            ' Now we will get the ">" before the NEXT PAGE
            t1 = InStr(ref1, text, ">")
            If t1 <> 0 Then
                ' Finally, extract what is in between and check if it says NEXT PAGE
                temp = Mid(text, t1 + 1, (ref2 - t1) - 1)
                If temp = "Next Page" Then
                    ' We'll want the URL for the page now too
                    temp = Mid(text, ref1, (t1 - ref1) - 1)
                    IsNextPage = temp
                    'Debug.Print temp
                End If
            End If
        End If
        ref1 = InStr(ref1 + 1, text, "/cgi-bin/HoTMaiL?")
'    End If
    Loop
End Function

Sub PS_MvFrm(frm As Form)
    'ReleaseCapture
    'Call SendMessage(frm.hwnd, &HA1, 2, 0&)
End Sub

Sub GetMessage(ByVal url As String)
    With frmhotmail
        msgURL = url
        .Socket.Action = SOCKET_DISCONNECT
        BatchNumber = 5
        .Socket.Action = 2
    End With
End Sub

Public Function ProcessMessage(ByVal text As String) As String
    ' We are now going to extract the body of a message from an HTML document
    ' The body of Hotmail messages are usually between <TT></TT> or <PRE></PRE> tags
    ' That makes this a whole lot simpler.
    
    Dim tt1 As Long, tt2 As Long
    Dim pre1 As Long, pre2 As Long
    Dim last As Long
    
    Dim msg As String
    Dim part1 As String, part2 As String
    Dim temp As String
    
    ' First, check for the <TT> tag
    tt1 = InStr(1, text, "<tt>")
    If tt1 <> 0 Then
        ' Now, find the </TT>
        tt2 = InStr(tt1, text, "</tt>")
        If tt2 <> 0 Then
            ' Finally, extract the message
            tt1 = tt1 + 4
            msg = Mid(text, tt1, tt2 - tt1)
            msg = "<HTML><HEAD><TITLE></TITLE></HEAD><BODY><FONT FACE='Courier New' SIZE=2>" & msg
            'msg = ProcessReturns(msg)
        End If
    Else
        ' There must be a <DIV> instead
        pre1 = InStr(1, text, "<div>", vbTextCompare)
        If pre1 <> 0 Then
            pre2 = InStr(pre1, text, "</div>")
            If pre2 <> 0 Then
                pre1 = pre1 + 5
                ' Before we cut it up, let's look for a <p> following the div.
                ' this would indicate an image or embedded attachment of sorts
                msg = RemoveCrap(Mid(text, pre1, pre2 - pre1))
                tt1 = InStr(pre1 - 2, text, "<p>")
                If tt1 <> 0 Then
                    ' Find the corresponding </p>
                    tt2 = InStr(tt1, text, "</p>")
                    If tt2 <> 0 Then
                        tt1 = tt1 + 3
                        If Left(Mid(text, tt1, tt2 - tt1), 4) = "<img" Then
                            temp = Mid(text, tt1, tt2 - tt1)
                            ' Extract the URL
                            tt1 = InStr(1, temp, Chr(34))
                            tt2 = InStr(tt1 + 1, temp, Chr(34))
                            tt1 = tt1 + 1
                            msg = msg & "<b><a href=" & Chr(34) & Mid(temp, tt1, tt2 - tt1) & Chr(34) & ">View Image</a></b>"
                            Messages(msgIDX).attach = True
                            Messages(msgIDX).attachURL = Mid(temp, tt1, tt2 - tt1)
                        End If
                    End If
                End If
                msg = "<HTML><HEAD><TITLE></TITLE></HEAD><BODY><FONT FACE='Courier New' SIZE=3>" & msg
            End If
        Else
            ' There must be a <PRE> instead
            pre1 = InStr(1, text, "<pre>")
            If pre1 <> 0 Then
                pre2 = InStr(pre1, text, "</pre>")
                If pre2 <> 0 Then
                    pre1 = pre1 + 5
                    msg = ProcessReturns(Mid(text, pre1, pre2 - pre1))
                    msg = "<HTML><HEAD><TITLE></TITLE></HEAD><BODY><FONT FACE='Courier New' SIZE=3>" & msg
                End If
            End If
        End If
    End If
    
    ' We can also check to see if there is an attachment with this message
    pre1 = InStr(1, text, "icon_clip.gif")
    If pre1 <> 0 Then
        pre2 = InStr(pre1, text, "<a")
        If pre2 <> 0 Then
            pre2 = pre2 + Len("<a href=") + 1
            tt1 = InStr(pre2, text, Chr(34))
            If tt1 <> 0 Then
                'tt1 = tt1 + Len("</a>")
                temp = Mid(text, pre2, tt1 - pre2)
                
                Messages(msgIDX).attach = True
                Messages(msgIDX).attachURL = temp
                
                msg = msg & "<P><B><a href=" & Chr(34) & "http://" & NewHost & temp & Chr(34) & ">View Attachment</a></B>"
            End If
        End If
    End If
    
    msg = msg & "</FONT></BODY></HTML>"
            
    ProcessMessage = msg
End Function

Public Function ProcessReturns(ByVal msg As String) As String
    Dim tt1 As Long
    Dim part1 As String, part2 As String
    Dim last As Long
    ' Now, we have to go through and replace carriage returns with <BR>
    tt1 = InStr(1, msg, Chr(10))
    Do Until tt1 = 0
        If tt1 <> last + Len("<BR>") Then
            part1 = Mid(msg, 1, tt1 - 1)
            part2 = Mid(msg, tt1 + 1, Len(msg) - tt1)
        
            msg = part1 & "<BR>" & part2
        Else
            part1 = Mid(msg, 1, tt1 - 1)
            part2 = Mid(msg, tt1 + 1, Len(msg) - tt1)
            
            msg = part1 & part2
        End If
        last = tt1 + 1
        tt1 = InStr(1, msg, Chr(10))
    Loop
    
    ProcessReturns = msg
End Function

Public Function RemoveCrap(ByVal msg As String) As String
    Dim tt1 As Long
    Dim part1 As String, part2 As String
    Dim last As Long
    ' Remove Crap Like these damn things: 
    tt1 = InStr(1, msg, Chr(11))
    Do Until tt1 = 0
        part1 = Mid(msg, 1, tt1 - 1)
        part2 = Mid(msg, tt1 + 1, Len(msg) - tt1)
            
        msg = part1 & part2
        
        last = tt1 + 1
        tt1 = InStr(1, msg, Chr(11))
    Loop
    RemoveCrap = msg
End Function

Public Function GetComposeURL(ByVal text As String) As String
    Dim c1 As Long, c2 As Long
    Dim temp As String
    
    ' The compose URL is hidden amongst the other functions found
    ' on the top bar on the hotmail page.  It's beside inbox and addresses
    ' We're going to look for part of the url, which won't ever change
    ' that's the key in all of this.  The part we want is "/cgi-bin/compose?a"
    
    c1 = InStr(1, text, "/cgi-bin/compose?")
    If c1 <> 0 Then
        ' Now we'll look for the ending quote
        c2 = InStr(c1, text, Chr(34))
        If c2 <> 0 Then
            ' and get what's between
            temp = Mid(text, c1, c2 - c1)
            GetComposeURL = temp
        End If
    End If
End Function

Public Sub ProcessComposePage(ByVal text As String)
    Dim c1 As Long, c2 As Long
    Dim b1 As Long, b2 As Long
    
    ' The two pieces of information we want are the URL that we
    ' post to, and the MsgHDrid.  What that is, I don't know.
    ' These two things are pretty simple to find.  The first, is
    ' right next to "action=" in a form.  The second is next to a
    ' "value=" in an hidden input box.  To be sure we get the URL,
    ' we're going to just look for "premail" and get the Id that follows
    
    ' Find the ID
    c1 = InStr(1, text, "premail/")
    If c1 <> 0 Then
        ' find the following quote
        c2 = InStr(c1, text, Chr(34))
        If c2 <> 0 Then
            ' Get the id
            sendURL = "/cgi-bin/" & Mid(text, c1, c2 - c1)
        End If
    End If
    
    ' Find the msghdrid
    c1 = InStr(1, text, "name=" & Chr(34) & "msghdrid" & Chr(34))
    If c1 <> 0 Then
        b1 = InStr(c1, text, "value=")
        If b1 <> 0 Then
            b1 = b1 + Len("value=") + 1
            b2 = InStr(b1, text, Chr(34))
            If b2 <> 0 Then
                msghdrid = Mid(text, b1, b2 - b1)
            End If
        End If
    End If
End Sub

Public Sub LoadAccounts()
    Dim filename As String
    Dim tAccount As HotmailAccount
    Dim fnum As Integer
    Dim recNum As Long
    filename = App.Path & "\Accounts.idx"
    If Dir(filename) <> "" Then
        addcount = 0
        ReDim Accounts(AccCount) As HotmailAccount
        recNum = 1
        ' Get the next availble file number
        fnum = FreeFile
        Open filename For Random Access Read As #fnum Len = Len(tAccount)
        Do Until EOF(fnum)
            ' Seek to the current record number
            Seek #fnum, recNum
            ' Get the address record from the file
            Get #fnum, , tAccount
            recNum = recNum + 1
            ' Make sure it's not a dud
            If tAccount.username <> "" And Left(tAccount.username, 1) <> Chr(0) Then
                ' Copy the information into the address array
                ReDim Preserve Accounts(AccCount) As HotmailAccount
                Accounts(AccCount) = tAccount
                AccCount = AccCount + 1
            End If
            
            If recNum * Len(tAccount) > LOF(fnum) Then
                Exit Do
            End If
        Loop
        Close #fnum
    End If
End Sub
Public Sub SaveAccounts()
    Dim filename As String
    Dim tAccount As HotmailAccount
    Dim fnum As Integer
    filename = App.Path & "\Accounts.idx"
    If Dir(filename) <> "" Then Kill (filename)
        ' Get the next availble file number
        fnum = FreeFile
        Open filename For Random Access Write As #fnum Len = Len(tAccount)
        For i = 0 To AccCount - 1
            ' Put the address to the file
            If RTrim(Accounts(i).username) <> "" Then
                Put #fnum, , Accounts(i)
            End If
        Next
        Close #fnum
End Sub
Public Sub AddAccount(ByVal username As String, ByVal loginname As String, ByVal password As String)
    ReDim Preserve Accounts(AccCount) As HotmailAccount
    With Accounts(AccCount)
        .loginname = loginname
        .username = username
        ' Encrypt this
        .password = password
    End With
    AccCount = AccCount + 1
End Sub

Function CheckVersion(ByVal newversion As String, ByVal oldversion As String) As Boolean
    Dim tNewVer As upVersion, tOldVer As upVersion
    Dim newer As Boolean
    
    ' Convert the string versions to structures
    tNewVer = ConvertToVersion(newversion)
    tOldVer = ConvertToVersion(oldversion)
    
    ' Compare the two structures
    If tNewVer.major > tOldVer.major Then
        newer = True
    ElseIf tNewVer.major = tOldVer.major Then
        ' Check the minor version
        If tNewVer.minor > tOldVer.minor Then
            newer = True
        Else
            ' Check the revision
            If tNewVer.rev > tOldVer.rev Then
                newer = True
            End If
        End If
    End If
    CheckVersion = newer
End Function

Function ConvertToVersion(ByVal ver As String) As upVersion
    Dim maj As Integer, min As Integer, rev As Integer
    Dim tVer As upVersion
    If ver <> "" Then
        ' Find first DOT
        min = InStr(1, ver, ".")
        If min Then
            tVer.major = Left(ver, min - 1)
            ' Find the next DOT
            rev = InStr(min + 1, ver, ".")
            If rev Then
                tVer.minor = Mid(ver, min + 1, (rev - 1) - min)
                tVer.rev = Mid(ver, rev + 1, Len(ver) - rev)
            Else
                tVer.minor = Mid(ver, min + 1, Len(ver) - min)
            End If
        Else
            tVer.major = ver
        End If
    End If
    ConvertToVersion = tVer
End Function
