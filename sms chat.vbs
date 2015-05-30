' Получаем значения переменных, которые передаются скрипту при его вызове
Number = WScript.Arguments(0)
Message = WScript.Arguments(1)
Message = Trim(Message)


' Обнуляем флаги, устанавливаем значение переменных
FlagVlr = 0
FlagHlr = 0
FlagSymbol=0
Users = ""
Count = 0


' Устанавливаем соответствие между переменными и именами файлов
vlr = "vlr.txt"
hlr = "hlr.txt"
vlr_back = "vlr_back.txt"


' Ищем номер абонента в списке HLR
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


' Ищем номер абонента в списке VLR
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


' Устанавливаем флаги в зависимости от наличия записи номера в файлах HLR и VLR

If (FlagHlr = 0) Then
Flaguser = 0
End If


If (FlagHlr = 1) And (FlagVlr = 0) Then
Flaguser = 1
End If


If (FlagHlr = 1) And (FlagVlr = 1) Then
Flaguser = 2
End If


' Устанавливаем значение текстовых сообщений

MessageAlreadyEnter = "Вы уже вошли в чат. Сейчас в группе: " & Users & "."
MessageNowInGroup = "Сейчас в группе: " & Users & "."
MessageEmpty = "Вы отправили пустое сообщение."


' Устанавливаем значение текстовых сообщений при пустой группе

If (Count = 0) Then
MessageNowInGroup = "В группе нет собеседников."
End If

If (Count = 1) Then
MessageAlreadyEnter = "Вы уже вошли в чат. В группе нет собеседников."
End If


' Обработка специальных сообщений

dlina = Len(Message)
symbol = Trim(Message)
symbol = Left(symbol, 1)
If (symbol = "#") Then
FlagSymbol = 1
End If


' Обработка запроса на вход в группу

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


' ======= Отправляем сообщение об успешном входе в группу =======

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "Вы успешно подключились к разговору. " + MessageNowInGroup
objMsg.Send()
End If


' Обработка запроса на вход в группу, если абонент уже находится в ней

If (Message = "#1") And (Flaguser = 2) Then
FlagVlr = 0
FlagSymbol = 0

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "Вы уже подключены к разговору. " + MessageNowInGroup
objMsg.Send()
End If


' Обработка повторного запроса на выход из группы

If (Message = "#0") And (Flaguser = 1) Then
FlagSymbol = 0
End If


' Обработка запроса на выход из группы

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


' ======= Отправляем сообщение об успешном выходе из группы =======

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "Вы успешно покинули группу."
objMsg.Send()

End If


' Обработка запроса на получение списка собеседников

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


' Обработка неверной или несуществующей команды

If (FlagSymbol = 1) And (FlagHlr = 1) Then
FlagVlr = 0
FlagSymbol = 0

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "Команда не распознана. Команды: #1 - вход, #0 - выход, #? - список собеседников."
objMsg.Send()

End If


' Обработка пустого сообщения

If (Message = "") And (FlagVlr = 1) Then
FlagVlr = 0

Set objSMSDriver = CreateObject("HeadwindGSM.SMSDriver")
objSMSDriver.Connect()
Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = MessageEmpty
objMsg.Send()
End If


' Обработка непустого сообщения

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

' ======= Отправляем квитанцию =======

Set objMsg = CreateObject("HeadwindGSM.SMSMessage")
objMsg.To = Number
objMsg.Body = "Message send to " & Count_receipt & " users. Text: '" & Message & "'"
objMsg.Send()

' ======= /Отправляем квитанцию =======


End If


' ======= Конец скрипта =======

WScript.Quit