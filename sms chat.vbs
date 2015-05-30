' �������� �������� ����������, ������� ���������� ������� ��� ��� ������
Number = WScript.Arguments(0)
Message = WScript.Arguments(1)
Message = Trim(Message)


' �������� �����, ������������� �������� ����������
FlagVlr = 0
FlagHlr = 0
FlagSymbol=0
Users = ""
Count = 0


' ������������� ������������ ����� ����������� � ������� ������
vlr = "vlr.txt"
hlr = "hlr.txt"
vlr_back = "vlr_back.txt"


' ���� ����� �������� � ������ HLR
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set filehlr = objFSO.OpenTextFile(hlr, 1)

Do Until filehlr.AtEndOfStream
    sLine = filehlr.ReadLine()
    nSpace = InStr(sLine, " ")
    
    If nSpace > 0 Then
    Number_hlr = Left(sLine, nSpace - 1)
    Name_hlr = Trim(Right(sLine, Len(sLine) - nSpace))
    End If  

    If (Number_hlr = Number) Then
    FlagHlr = 1
    Name = Name_hlr
    End If
    
    If (Len(Users) = 2) Then
    Users = ""
    End If
Loop
filehlr.Close


' ���� ����� �������� � ������ VLR
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set filevlr = objFSO.OpenTextFile(vlr, 1)

Do Until filevlr.AtEndOfStream
    sLine = filevlr.ReadLine()
    nSpace = InStr(sLine, " ")
    
    If nSpace > 0 Then
    Users = Users & ", "
    Number_vlr = Left(sLine, nSpace - 1)
    Name_vlr = Trim(Right(sLine, Len(sLine) - nSpace))
    End If  

    Count = Count + 1

    If (Number_vlr = Number) Then
    FlagVlr = 1
    Name = Name_vlr
    End If

    
    If (Len(Users) = 2) Then
    Users = ""
    End If

    Users = Users & Name_vlr
Loop
filevlr.Close


' ������������� ����� � ����������� �� ������� ������ ������ � ������ HLR � VLR

If (FlagHlr = 0) Then
Flaguser = 0
End If


If (FlagHlr = 1) And (FlagVlr = 0) Then
Flaguser = 1
End If


If (FlagHlr = 1) And (FlagVlr = 1) Then
Flaguser = 2
End If


' ������������� �������� ��������� ���������

MessageAlreadyEnter = "�� ��� ����� � ���. ������ � ������: " & Users & "."
MessageNowInGroup = "������ � ������: " & Users & "."
MessageEmpty = "�� ��������� ������ ���������."


' ������������� �������� ��������� ��������� ��� ������ ������

If (Count = 0) Then
MessageNowInGroup = "� ������ ��� ������������."
End If

If (Count = 1) Then
MessageAlreadyEnter = "�� ��� ����� � ���. � ������ ��� ������������."
End If


' ��������� ����������� ���������

dlina = Len(Message)
symbol = Trim(Message)
symbol = Left(symbol, 1)
If (symbol = "#") Then
FlagSymbol = 1
End If


' ��������� ������� �� ���� � ������

If (Message = "#1") And (Flaguser = 1) Then
FlagVlr = 0
FlagSymbol = 0

Set filevlr = objFSO.OpenTextFile(vlr, 1)
Set filevlrback = objFSO.OpenTextFile(vlr_back, 2)
rec = Number + " " + Name
filevlrback.WriteLine(rec)

Do Until filevlr.AtEndOfStream
    sw = filevlr.ReadLine()
    filevlrback.WriteLine(sw)
    Loop

filevlr.Close
filevlrback.Close

Set filevlr = objFSO.OpenTextFile(vlr, 2)
Set filevlrback = objFSO.OpenTextFile(vlr_back, 1)
Do Until filevlrback.AtEndOfStream
    sw = filevlrback.ReadLine()
    filevlr.WriteLine(sw)
    Loop

filevlr.Close
filevlrback.Close


' ======= ���������� ��������� �� �������� ����� � ������ =======

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "�� ������� ������������ � ���������. " + MessageNowInGroup
objMsg.Send()
End If


' ��������� ������� �� ���� � ������, ���� ������� ��� ��������� � ���

If (Message = "#1") And (Flaguser = 2) Then
FlagVlr = 0
FlagSymbol = 0

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "�� ��� ���������� � ���������. " + MessageNowInGroup
objMsg.Send()
End If


' ��������� ���������� ������� �� ����� �� ������

If (Message = "#0") And (Flaguser = 1) Then
FlagSymbol = 0
End If


' ��������� ������� �� ����� �� ������

If (Message = "#0") And (FlagVlr = 1) Then
FlagVlr = 0
FlagSymbol = 0


Set filevlr = objFSO.OpenTextFile(vlr, 1)
Set filevlrback = objFSO.OpenTextFile(vlr_back, 2)

Do Until filevlr.AtEndOfStream
    sLine = filevlr.ReadLine()
    nSpace = InStr(sLine, " ")
    
    If nSpace > 0 Then
    Number_vlr = Left(sLine, nSpace - 1)
    Name_vlr = Trim(Right(sLine, Len(sLine) - nSpace))
    rec = Number_vlr + " " + Name_vlr
    End If  

    If (Number_vlr <> Number) Then
    filevlrback.WriteLine(rec)
    End If

    

Loop

filevlr.Close
filevlrback.Close

Set filevlr = objFSO.OpenTextFile(vlr, 2)
Set filevlrback = objFSO.OpenTextFile(vlr_back, 1)
Do Until filevlrback.AtEndOfStream
rec = filevlrback.ReadLine()
filevlr.WriteLine(rec)
Loop
filevlr.Close
filevlrback.Close


' ======= ���������� ��������� �� �������� ������ �� ������ =======

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "�� ������� �������� ������."
objMsg.Send()

End If


' ��������� ������� �� ��������� ������ ������������

If (Message = "#?") And (FlagVlr = 1) Then
FlagVlr = 0
FlagSymbol = 0

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = MessageNowInGroup
objMsg.Send()
End If


' ��������� �������� ��� �������������� �������

If (FlagSymbol = 1) And (FlagHlr = 1) Then
FlagVlr = 0
FlagSymbol = 0

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "������� �� ����������. �������: #1 - ����, #0 - �����, #? - ������ ������������."
objMsg.Send()

End If


' ��������� ������� ���������

If (Message = "") And (FlagVlr = 1) Then
FlagVlr = 0

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = MessageEmpty
objMsg.Send()
End If


' ��������� ��������� ���������

If (FlagVlr = 1) Then
        Message = Trim(Message)
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set filevlr = objFSO.OpenTextFile(vlr, 1)
        Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
        objSMSDriver.Connect()
        Do Until filevlr.AtEndOfStream
        sLine = filevlr.ReadLine
        nSpace = InStr(sLine, " ")
            If nSpace > 0 Then
            Number_buffer = Left(sLine, nSpace - 1)
            Name_buffer = Trim(Right(sLine, Len(sLine) - nSpace))
                If (Number_buffer <> Number) Then
                Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
                objMsg.To = Number_buffer
                objMsg.Body = Name & ": " & Message
                objMsg.Send()
                End If
            End If
Loop
filevlr.Close

Count_receipt = Count - 1

' ======= ���������� ��������� =======

Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "Message send to " & Count_receipt & " users. Text: '" & Message & "'"
objMsg.Send()

' ======= /���������� ��������� =======


End If


' ======= ����� ������� =======

WScript.Quit