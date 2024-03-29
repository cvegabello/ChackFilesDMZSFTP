﻿
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Net
Imports Microsoft.Win32
Imports System.Net.Mail



Module Utils
    Public appName As String = "CheckFiles"

    Public Function GetSetting(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Return Key.GetValue(Keyname, DefVal)


        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSetting(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Sub SalvarSettingConfiServerDB(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Function GetSettingConfigServerDB(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSettingEmail(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey("eMail")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub
    Public Function GetSettingEmail(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey("eMail")
            Return Key.GetValue(Keyname, DefVal)

        Catch
            Return DefVal
        End Try

    End Function

    Public Sub SalvarSettingEmailOnePlace(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey("eMailOnePlace")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub
    Public Function GetSettingEmailOnePlace(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey("eMailOnePlace")
            Return Key.GetValue(Keyname, DefVal)

        Catch
            Return DefVal
        End Try

    End Function




    Function returnFecha(ByVal toKnowCdc As Integer) As Date
        Dim initDate As Date
        initDate = "1/1/1986"

        Today = System.DateTime.Now
        returnFecha = initDate.AddDays(toKnowCdc)

        'returnFecha = (initDate + toKnowCdc) - 1

    End Function


    Function convertDateInt(ByVal toKnowDate As Date) As Integer
        Dim initDate As Date
        Dim diff1 As TimeSpan
        initDate = "1/1/1900"

        diff1 = toKnowDate.Subtract(initDate)
        convertDateInt = diff1.Days + 2

    End Function

    Function returnCDC(ByVal toKnowDate As Date) As Integer
        Dim initDate As Date
        Dim diff1 As TimeSpan
        initDate = "1/1/1986"

        diff1 = toKnowDate.Subtract(initDate)
        returnCDC = diff1.Days + 1

    End Function
    Function returnCDC_OTG(ByVal toKnowDate As Date) As Integer
        Dim initDate As Date
        Dim diff1 As TimeSpan
        initDate = "5/1/2013"

        diff1 = toKnowDate.Subtract(initDate)
        returnCDC_OTG = diff1.Days + 1

    End Function

    Function returnCDC_OP(ByVal toKnowDate As Date) As Integer
        'Dim initDate As Date
        ''Dim diff1 As TimeSpan
        'initDate = "5/1/2013"
        returnCDC_OP = returnCDC(toKnowDate) - 3652
        'diff1 = toKnowDate.Subtract(initDate)
        'returnCDC_OP = diff1.Days + 1

    End Function

    Function returnDayWeek(ByVal dayWeek As Integer) As String

        Select Case dayWeek
            Case 1
                returnDayWeek = "SUNDAY"

            Case 2
                returnDayWeek = "MONDAY"

            Case 3
                returnDayWeek = "TUESDAY"

            Case 4
                returnDayWeek = "WEDNESDAY"

            Case 5
                returnDayWeek = "THURSDAY"

            Case 6
                returnDayWeek = "FRIDAY"

            Case 7
                returnDayWeek = "SATURDAY"

            Case Else

                returnDayWeek = ""

        End Select

    End Function

    Function getInfoFileTbl(ByVal toKnowDate As Date) As DataTable


        Dim filesLocationTbl As New DataTable("FILES_LOCATION")
        Dim nuevaFila As DataRow
        Dim reader1 As StreamReader
        Dim strLine, strValue As String
        Dim substrings() As String

        'cdcInt = returnCDC(toKnowDate)
        filesLocationTbl = Utils.Despliegue

        strValue = ""
        Try
            reader1 = New StreamReader("listFilesExistTest.txt", Encoding.UTF7)
            strLine = ""

            Do While Not (strLine Is Nothing)
                strLine = reader1.ReadLine()
                If Not (strLine Is Nothing) Then
                    substrings = strLine.Split("|")
                    nuevaFila = filesLocationTbl.NewRow()
                    nuevaFila(0) = substrings(0)
                    nuevaFila(1) = substrings(1)
                    filesLocationTbl.Rows.Add(nuevaFila)
                End If
            Loop
            reader1.Close()

        Catch ex As Exception
            MsgBox("Error reading file. " & ex.Message, MsgBoxStyle.Information, "Error")

        End Try

        Return filesLocationTbl


    End Function

    Function Despliegue() As DataTable
        Dim fileName As New DataColumn
        Dim location As New DataColumn
        Dim exist As New DataColumn
        'Dim countDown As New DataColumn
        Dim checkFilesTbl As New DataTable("FILES_LOCATION")

        fileName.ColumnName = "FILE_NAME"
        fileName.DataType = System.Type.GetType("System.String")

        location.ColumnName = "LOCATION"
        location.DataType = System.Type.GetType("System.String")
        'location.MaxLength = 50

        exist.ColumnName = "NUM_DRAW2"
        exist.DataType = System.Type.GetType("System.String")
        'exist.MaxLength = 50


        checkFilesTbl.Columns.Add(fileName)
        checkFilesTbl.Columns.Add(location)
        checkFilesTbl.Columns.Add(exist)
        'drawsGamesTbl.Columns.Add(countDown)

        Despliegue = checkFilesTbl

    End Function

    Function sent_Email(ByVal smtpHostStr As String, ByVal toStr As String, ByVal attachStr As String, ByVal subjectStr As String, ByVal bodyStr As String) As Integer
        Dim SMTPserver As New SmtpClient
        Dim mail As New MailMessage
        Dim oAttch As Attachment = New Attachment(attachStr)


        Try
            SMTPserver.Host = smtpHostStr
            mail = New MailMessage
            mail.From = New MailAddress("do.not.reply@gtech-noreply.com")
            mail.To.Add(toStr) 'The Man you want to send the message to him
            mail.Subject = subjectStr
            mail.Body = bodyStr
            mail.Attachments.Add(oAttch)
            SMTPserver.Send(mail)
            oAttch.Dispose()
            Return 0
            'MessageBox.Show("Done!", "Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Error)
            oAttch.Dispose()
            Return -1

        End Try


    End Function



    'Function DespliegueDateInfo() As DataTable
    '    Dim titleDateCol As New DataColumn
    '    Dim infoDateCol As New DataColumn
    '    Dim dateTbl As New DataTable("DATEINFO")

    '    titleDateCol.ColumnName = "TITLE_DATE"
    '    titleDateCol.DataType = System.Type.GetType("System.String")

    '    infoDateCol.ColumnName = "INFO_DATE"
    '    infoDateCol.DataType = System.Type.GetType("System.String")
    '    infoDateCol.MaxLength = 50

    '    dateTbl.Columns.Add(titleDateCol)
    '    dateTbl.Columns.Add(infoDateCol)

    '    DespliegueDateInfo = dateTbl

    'End Function


    'Function DespliegueTake5Winners() As DataTable
    '    Dim dayName As New DataColumn
    '    Dim cdcInt As New DataColumn
    '    Dim numDraw As New DataColumn
    '    Dim winners As New DataColumn
    '    Dim winnersTake5Tbl As New DataTable("TAKE5_WINNERS")

    '    dayName.ColumnName = "DAY"
    '    dayName.DataType = System.Type.GetType("System.String")

    '    cdcInt.ColumnName = "CDC"
    '    cdcInt.DataType = System.Type.GetType("System.Single")


    '    'numDraw.ColumnName = "NUM_DRAW"
    '    'numDraw.DataType = System.Type.GetType("System.Single")

    '    winners.ColumnName = "WINNERS"
    '    winners.DataType = System.Type.GetType("System.Single")

    '    winnersTake5Tbl.Columns.Add(dayName)
    '    winnersTake5Tbl.Columns.Add(cdcInt)
    '    'winnersTake5Tbl.Columns.Add(numDraw)
    '    winnersTake5Tbl.Columns.Add(winners)

    '    DespliegueTake5Winners = winnersTake5Tbl

    'End Function


    'Sub getAndSaveWinnersALLProduct(ByVal toKnowDate As DateTime, ByRef errCodeInt As Integer, ByRef errMessageStr As String)

    '    Dim tableQproducts As DataTable
    '    Dim row1 As DataRow
    '    Dim idProductInt, productSysCodeInt, drawNumberInt, cdcInt, stateCodInt, numWinnersInt, numrecordsInt, errCodInt As Integer
    '    Dim oldJackPotDbl, newJackPotDbl As Double
    '    Dim productSysNameStr, numWinnerStr, nameDayWeekStr, msgErrorStr, strFilePath As String
    '    Dim parameters1(4) As MySql.Data.MySqlClient.MySqlParameter
    '    Dim parameters2(1) As MySql.Data.MySqlClient.MySqlParameter
    '    Dim table1 As DataTable
    '    Dim resBool As Boolean
    '    'Dim errCodeInt As Integer = 0
    '    'Dim errMessageStr As String = ""

    '    Dim draw(3) As String


    '    tableQproducts = dataSourceAccess.RunStoreQueryWithoutParameters("QProductsWinners")

    '    For Each row1 In tableQproducts.Rows
    '        productSysCodeInt = row1("ProductSysCode")
    '        productSysNameStr = row1("productSysName")
    '        idProductInt = row1("IdProduct")
    '        If (IsDBNull(row1("IdProduct"))) Then
    '            oldJackPotDbl = 0
    '        Else
    '            oldJackPotDbl = row1("currentJackpot")
    '        End If


    '        If IsDrawDay(productSysCodeInt, toKnowDate) Then
    '            numWinnerStr = Trim(getWinnersXProduct(toKnowDate, productSysCodeInt, errCodeInt, errMessageStr))

    '            If errCodeInt = 0 Then
    '                drawNumberInt = getDrawNumberXProductXday(productSysCodeInt, toKnowDate, 0)
    '                drawNumberInt -= 1
    '                cdcInt = returnCDC(toKnowDate)
    '                cdcInt -= 1

    '                stateCodInt = 35
    '                numWinnersInt = CInt(numWinnerStr)
    '                If numWinnersInt > 0 Then
    '                    parameters1(0) = New MySql.Data.MySqlClient.MySqlParameter("@codproduct", idProductInt)
    '                    parameters1(1) = New MySql.Data.MySqlClient.MySqlParameter("@cdc", cdcInt)
    '                    parameters1(2) = New MySql.Data.MySqlClient.MySqlParameter("@codstate", stateCodInt)
    '                    parameters1(3) = New MySql.Data.MySqlClient.MySqlParameter("@numwinners", numWinnersInt)
    '                    parameters1(4) = New MySql.Data.MySqlClient.MySqlParameter("@dateTimeInsert", Now)
    '                    msgErrorStr = ""
    '                    dataSourceAccess.RunStoreProcedureParametersNonQuery("insertWinnnerXstate", parameters1, msgErrorStr)
    '                End If

    '                parameters2(0) = New MySql.Data.MySqlClient.MySqlParameter("@codproduct", idProductInt)
    '                parameters2(1) = New MySql.Data.MySqlClient.MySqlParameter("@cdc", cdcInt)
    '                table1 = dataSourceAccess.RunStructureProcedureReturnDtable("returnTotalWinners", parameters2)
    '                numrecordsInt = table1.Rows.Count

    '                If numrecordsInt = 0 Then
    '                    numWinnerStr = "0"
    '                Else
    '                    numWinnerStr = CStr(table1.Rows(0).Item(2))
    '                End If

    '                nameDayWeekStr = WeekdayName(Weekday(toKnowDate.AddDays(-1)))

    '                If productSysCodeInt = 8 Or productSysCodeInt = 12 Or productSysCodeInt = 15 Then

    '                    newJackPotDbl = readJackpotFromFile(productSysCodeInt, toKnowDate, errCodeInt, errMessageStr)
    '                    If errCodeInt <> 0 Then
    '                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
    '                    End If
    '                Else
    '                    newJackPotDbl = 0
    '                End If

    '                'newJackPotDbl = leerJackpotFromPaginaWeb("http://www.lotteryusa.com/lottery/NY/NY_fcur.html", productSysCodeInt)

    '                If (newJackPotDbl <> oldJackPotDbl) Or (newJackPotDbl = 0) Then
    '                    msgErrorStr = ""
    '                    errCodInt = deleteHistoryDaily(idProductInt, cdcInt, msgErrorStr)
    '                    If errCodInt = 0 Then
    '                        saveHistoryDaily(drawNumberInt, cdcInt, productSysNameStr, idProductInt, numWinnerStr, oldJackPotDbl, nameDayWeekStr)

    '                        updateJackPot(productSysCodeInt, newJackPotDbl)

    '                    Else
    '                        MsgBox("Error in delete HistoryDaily: " & msgErrorStr, MsgBoxStyle.Critical, "Error")
    '                    End If

    '                End If

    '            Else
    '                Exit Sub

    '            End If


    '        End If
    '    Next

    '    strFilePath = Utils.GetSetting(appName, "UbiWinFileLocation", "").ToString()
    '    msgErrorStr = ""
    '    resBool = moveFiles(strFilePath, strFilePath & "\History", toKnowDate, msgErrorStr)

    '    If Not resBool Then
    '        MsgBox("Error in move the files to History folder: " & msgErrorStr, MsgBoxStyle.Critical, "Error")

    '    End If



    'End Sub

    'Function readJackpotFromFile(ByVal productSysCodeInt As Integer, ByVal toKnowDate As Date, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As Double

    '    Dim reader1 As StreamReader
    '    Dim strLinea As String
    '    Dim strFile As String = ""
    '    Dim drawNumberStr, strValor, strFilePath As String
    '    Dim winnerStr As String = ""
    '    Dim drawNumberInt As Integer
    '    Dim jackPotDbl As Double

    '    strFilePath = Utils.GetSetting(appName, "UbiWinFileLocation", "").ToString()


    '    drawNumberInt = getDrawNumberXProductXday(productSysCodeInt, toKnowDate, 0)
    '    drawNumberStr = Trim(Str(drawNumberInt))
    '    drawNumberStr = drawNumberStr.PadLeft(6, "0")

    '    Select Case productSysCodeInt
    '        Case 8
    '            strFile = strFilePath & "\jackpot_" & drawNumberStr & "_loto.txt"

    '        Case 12
    '            strFile = strFilePath & "\jackpot_" & drawNumberStr & "_bigg.txt"

    '        Case 15
    '            strFile = strFilePath & "\jackpot_" & drawNumberStr & "_pwrb.txt"

    '    End Select

    '    Try
    '        reader1 = New StreamReader(strFile, Encoding.UTF7)
    '        strLinea = ""
    '        strLinea = reader1.ReadLine()
    '        strValor = Mid(strLinea, 41, 20)
    '        jackPotDbl = CDbl(strValor)
    '        reader1.Close()
    '        errCodeInt = 0
    '        Return jackpotDbl

    '    Catch ex As Exception

    '        errCodeInt = -1
    '        errMessageStr = ex.Message

    '    End Try



    'End Function

    'Function moveFiles(ByVal fromLocation As String, ByVal ToLocation As String, ByVal fecha As Date, ByRef msgError As String) As Boolean
    '    Dim nameFoundFileStr As String
    '    Dim monthStr, dayStr, yearStr As String

    '    monthStr = CStr(fecha.Month).PadLeft(2, "0")
    '    dayStr = CStr(fecha.Day).PadLeft(2, "0")
    '    yearStr = CStr(fecha.Year).PadLeft(4, "0")

    '    For Each foundFile As String In My.Computer.FileSystem.GetFiles(fromLocation)
    '        nameFoundFileStr = My.Computer.FileSystem.GetName(foundFile)
    '        If nameFoundFileStr <> "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt" And nameFoundFileStr <> "hostMode.txt" Then
    '            Try
    '                My.Computer.FileSystem.MoveFile(foundFile, ToLocation & "\" & nameFoundFileStr, True)
    '            Catch ex As Exception
    '                msgError = ex.Message
    '                Return False
    '            End Try
    '        End If
    '    Next
    '    Return True

    'End Function












    ''Sub getAndSaveWinnersALLProduct(ByVal toKnowDate As DateTime)

    ''    Dim tableQproducts As DataTable
    ''    Dim row1 As DataRow
    ''    Dim idProductInt, productSysCodeInt, drawNumberInt, cdcInt, stateCodInt, numWinnersInt, numrecordsInt, errCodInt As Integer
    ''    Dim oldJackPotDbl, newJackPotDbl As Double
    ''    Dim productSysNameStr, numWinnerStr, nameDayWeekStr, msgErrorStr As String
    ''    Dim parameters1(4) As MySql.Data.MySqlClient.MySqlParameter
    ''    Dim parameters2(1) As MySql.Data.MySqlClient.MySqlParameter
    ''    Dim table1 As DataTable

    ''    Dim draw(3) As String


    ''    tableQproducts = dataSourceAccess.RunStoreQueryWithoutParameters("QProductsWinners")

    ''    For Each row1 In tableQproducts.Rows
    ''        productSysCodeInt = row1("ProductSysCode")
    ''        productSysNameStr = row1("productSysName")
    ''        idProductInt = row1("IdProduct")
    ''        If (IsDBNull(row1("IdProduct"))) Then
    ''            oldJackPotDbl = 0
    ''        Else
    ''            oldJackPotDbl = row1("currentJackpot")
    ''        End If


    ''        If IsDrawDay(productSysCodeInt, toKnowDate) Then
    ''            numWinnerStr = Trim(getWinnersXProduct(toKnowDate, productSysCodeInt))
    ''            drawNumberInt = getDrawNumberXProductXday(productSysCodeInt, toKnowDate, 0)
    ''            drawNumberInt -= 1
    ''            cdcInt = returnCDC(toKnowDate)
    ''            cdcInt -= 1

    ''            stateCodInt = 35
    ''            numWinnersInt = CInt(numWinnerStr)
    ''            If numWinnersInt > 0 Then
    ''                parameters1(0) = New MySql.Data.MySqlClient.MySqlParameter("@codproduct", idProductInt)
    ''                parameters1(1) = New MySql.Data.MySqlClient.MySqlParameter("@cdc", cdcInt)
    ''                parameters1(2) = New MySql.Data.MySqlClient.MySqlParameter("@codstate", stateCodInt)
    ''                parameters1(3) = New MySql.Data.MySqlClient.MySqlParameter("@numwinners", numWinnersInt)
    ''                parameters1(4) = New MySql.Data.MySqlClient.MySqlParameter("@dateTimeInsert", Now)
    ''                msgErrorStr = ""
    ''                dataSourceAccess.RunStoreProcedureParametersNonQuery("insertWinnnerXstate", parameters1, msgErrorStr)
    ''            End If

    ''            parameters2(0) = New MySql.Data.MySqlClient.MySqlParameter("@codproduct", idProductInt)
    ''            parameters2(1) = New MySql.Data.MySqlClient.MySqlParameter("@cdc", cdcInt)
    ''            table1 = dataSourceAccess.RunStructureProcedureReturnDtable("returnTotalWinners", parameters2)
    ''            numrecordsInt = table1.Rows.Count

    ''            If numrecordsInt = 0 Then
    ''                numWinnerStr = "0"
    ''            Else
    ''                numWinnerStr = CStr(table1.Rows(0).Item(2))
    ''            End If

    ''            nameDayWeekStr = WeekdayName(Weekday(toKnowDate.AddDays(-1)))

    ''            newJackPotDbl = readJackpotFromPaginaWeb("http://www.lotteryusa.com/lottery/NY/NY_fcur.html", productSysCodeInt)

    ''            If (newJackPotDbl <> oldJackPotDbl) Or (newJackPotDbl = 0) Then
    ''                msgErrorStr = ""
    ''                errCodInt = deleteHistoryDaily(idProductInt, cdcInt, msgErrorStr)
    ''                If errCodInt = 0 Then
    ''                    saveHistoryDaily(drawNumberInt, cdcInt, productSysNameStr, idProductInt, numWinnerStr, oldJackPotDbl, nameDayWeekStr)

    ''                    updateJackPot(productSysCodeInt, newJackPotDbl)

    ''                Else
    ''                    MsgBox("Error in delete HistoryDaily: " & msgErrorStr, MsgBoxStyle.Critical, "Error")
    ''                End If


    ''            End If



    ''        End If
    ''    Next
    ''End Sub

    'Function getDrawNumberXProductXday(ByVal codSysProduct As Integer, ByVal toKnowDate As DateTime, ByVal dayDraw As Integer) As Integer

    '    Dim num1 As Integer
    '    Dim num3 As Integer
    '    Dim dayWeekInt As Integer
    '    Dim moduloRes As Integer
    '    Dim cocienteRes As Integer

    '    num1 = convertDateInt(toKnowDate)
    '    dayWeekInt = Weekday(toKnowDate)

    '    Select Case codSysProduct
    '        Case 8 'Loto

    '            Select Case dayWeekInt
    '                Case 5
    '                    moduloRes = (num1 + 2) Mod 7
    '                    cocienteRes = ((num1 + 2) \ 7) * 2
    '                Case 6
    '                    moduloRes = (num1 + 1) Mod 7
    '                    cocienteRes = ((num1 + 1) \ 7) * 2
    '                Case Else
    '                    moduloRes = num1 Mod 7
    '                    cocienteRes = (num1 \ 7) * 2

    '            End Select
    '            If (moduloRes < 1) Then
    '                num3 = cocienteRes - 11488
    '            Else
    '                num3 = cocienteRes - 11487
    '            End If
    '            num3 += 2357



    '        Case 9 'PCK3 -NUMBERS-

    '            Select Case dayDraw
    '                Case 1 'Midday
    '                    num3 = ((2 * num1) - 69216)
    '                Case 2 'Evenning
    '                    num3 = (((2 * num1) - 69216) + 1)
    '            End Select

    '        Case 10
    '            num3 = (num1 - 35525)

    '        Case 11

    '        Case 12 'BIGG
    '            If ((num1 Mod 7) < 4) Then
    '                num3 = (((num1 \ 7) * 2) - 10682)
    '            Else
    '                num3 = (((num1 \ 7) * 2) - 10681)
    '            End If

    '        Case 13 'LIFE

    '            Select Case dayWeekInt
    '                Case 6
    '                    num3 = (((num1 \ 7) * 2) - 11941)

    '                Case Else
    '                    If ((num1 Mod 7) < 3) Then
    '                        num3 = (((num1 \ 7) * 2) - 11943)
    '                    Else
    '                        num3 = (((num1 \ 7) * 2) - 11942)
    '                    End If

    '            End Select


    '        Case 14

    '            Select Case dayDraw
    '                Case 1 'Midday
    '                    num3 = ((2 * num1) - 69216)
    '                Case 2 'Evenning
    '                    num3 = (((2 * num1) - 69216) + 1)
    '            End Select


    '        Case 15 'PowerBall

    '            Select Case dayWeekInt
    '                Case 5
    '                    moduloRes = (num1 + 2) Mod 7
    '                    cocienteRes = ((num1 + 2) \ 7) * 2
    '                Case 6
    '                    moduloRes = (num1 + 1) Mod 7
    '                    cocienteRes = ((num1 + 1) \ 7) * 2
    '                Case Else
    '                    moduloRes = num1 Mod 7
    '                    cocienteRes = (num1 \ 7) * 2

    '            End Select
    '            If (moduloRes < 1) Then
    '                num3 = cocienteRes - 11488
    '            Else
    '                num3 = cocienteRes - 11487
    '            End If

    '        Case 27 'DKNO
    '            num3 = (num1 - 31979)

    '    End Select
    '    Return num3
    'End Function

    'Function getInfoDateTbl(ByVal toKnowDate As String) As DataTable

    '    Dim drawGamesTbl As New DataTable("DATEINFO")
    '    Dim nuevaFila As DataRow

    '    drawGamesTbl = Utils.DespliegueDateInfo()

    '    nuevaFila = drawGamesTbl.NewRow()
    '    drawGamesTbl.Rows.Add(nuevaFila)
    '    nuevaFila(0) = "DATE:"
    '    nuevaFila(1) = toKnowDate


    '    nuevaFila = drawGamesTbl.NewRow()
    '    drawGamesTbl.Rows.Add(nuevaFila)
    '    nuevaFila(0) = "CDC:"
    '    'currentTimeStr = Format(TimeOfDay, "HH:mm:ss")
    '    If TimeOfDay >= "00:00:01" And TimeOfDay < "03:30:00" Then
    '        nuevaFila(1) = returnCDC(toKnowDate) - 1
    '    Else
    '        nuevaFila(1) = returnCDC(toKnowDate)
    '    End If


    '    nuevaFila = drawGamesTbl.NewRow()
    '    drawGamesTbl.Rows.Add(nuevaFila)
    '    nuevaFila(0) = "DAY:"
    '    nuevaFila(1) = returnDayWeek(Weekday(toKnowDate))


    '    Return drawGamesTbl
    'End Function



    'Function getWinnersXProduct(ByVal toKnowDate As DateTime, ByVal codSysProduct As Integer, ByRef errorCode As Integer, ByRef errorMessage As String) As String

    '    Dim reader1 As StreamReader
    '    Dim reader2 As XmlReader
    '    Dim strLinea As String
    '    Dim strFile, drawNumberStr, strValor, drawNumberPrevStr, strFileOld, strFilePath As String
    '    Dim winnerStr As String = ""
    '    Dim drawNumberInt As Integer

    '    drawNumberInt = Utils.getDrawNumberXProductXday(codSysProduct, toKnowDate, 0)
    '    drawNumberInt -= 1
    '    drawNumberPrevStr = Trim(Str(drawNumberInt - 1))
    '    drawNumberStr = Trim(Str(drawNumberInt))
    '    strFilePath = Utils.GetSetting(appName, "UbiWinFileLocation", "").ToString()

    '    'strFile = Application.StartupPath
    '    ChDir(strFilePath)

    '    Select Case codSysProduct
    '        Case 10 'CSH5
    '            drawNumberStr = drawNumberStr.PadLeft(6, "0")
    '            drawNumberPrevStr = drawNumberPrevStr.PadLeft(6, "0")

    '            strFileOld = strFilePath & "\winner_summary_report_p010_d" & drawNumberPrevStr & "_wincnt_english.rep"
    '            Try
    '                FileSystem.Kill(strFileOld)
    '            Catch ex As Exception

    '            End Try

    '            strFile = strFilePath & "\winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep"

    '            Try
    '                reader1 = New StreamReader(strFile, Encoding.UTF7)
    '                strLinea = ""

    '                Do While Not (strLinea Is Nothing)
    '                    strLinea = reader1.ReadLine()
    '                    strValor = Mid(strLinea, 3, 3)
    '                    If strValor = "5/5" Then
    '                        winnerStr = Mid(strLinea, 25, 6)
    '                        Exit Do
    '                    End If

    '                Loop
    '                reader1.Close()
    '            Catch ex As Exception

    '                errorCode = -1
    '                errorMessage = "The file TAKE5 WINNER SUMMARY -winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep-, is not in " & strFilePath
    '                winnerStr = "0"

    '            End Try



    '        Case 8 'LOTTO
    '            drawNumberStr = drawNumberStr.PadLeft(6, "0")
    '            drawNumberPrevStr = drawNumberPrevStr.PadLeft(6, "0")

    '            strFileOld = strFilePath & "\winner_summary_report_p008_d" & drawNumberPrevStr & "_wincnt_english.rep"
    '            Try
    '                FileSystem.Kill(strFileOld)
    '            Catch ex As Exception

    '            End Try

    '            strFile = strFilePath & "\winner_summary_report_p008_d" & drawNumberStr & "_wincnt_english.rep"

    '            Try
    '                reader1 = New StreamReader(strFile, Encoding.UTF7)
    '                strLinea = ""
    '                Do While Not (strLinea Is Nothing)
    '                    strLinea = reader1.ReadLine()
    '                    strValor = Mid(strLinea, 3, 3)
    '                    If strValor = "6/6" Then
    '                        winnerStr = Mid(strLinea, 25, 6)
    '                        Exit Do
    '                    End If
    '                Loop
    '                reader1.Close()

    '            Catch ex As Exception

    '                errorCode = -1
    '                errorMessage = "The file LOTTO WINNER SUMMARY -winner_summary_report_p008_d" & drawNumberStr & "_wincnt_english.rep-, is not in " & strFilePath
    '                winnerStr = "0"

    '            End Try




    '        Case 12 'BIGG
    '            drawNumberStr = drawNumberStr.PadLeft(4, "0")
    '            drawNumberPrevStr = drawNumberPrevStr.PadLeft(4, "0")

    '            strFileOld = strFilePath & (("\winners_mm_" & drawNumberPrevStr) & ".xml")
    '            Try
    '                FileSystem.Kill(strFileOld)
    '            Catch ex As Exception

    '            End Try


    '            strFile = strFilePath & (("\winners_mm_" & drawNumberStr) & ".xml")

    '            Try
    '                reader2 = XmlReader.Create(strFile)
    '                Do While reader2.Read()
    '                    If (reader2.NodeType = XmlNodeType.Element AndAlso reader2.Name = "winners") Then
    '                        winnerStr = reader2.GetAttribute(3)
    '                    End If
    '                Loop
    '                reader2.Close()
    '            Catch ex As Exception
    '                errorCode = -1
    '                errorMessage = "The file BIGG WINNERS  -winners_mm_" & drawNumberStr & ".xml- is not in " & strFilePath
    '                winnerStr = "0"

    '            End Try



    '        Case 15 'PWRB
    '            drawNumberStr = drawNumberStr.PadLeft(4, "0")
    '            drawNumberPrevStr = drawNumberPrevStr.PadLeft(4, "0")

    '            strFileOld = strFilePath & (("\winners_pb_" & drawNumberPrevStr) & ".xml")
    '            Try
    '                FileSystem.Kill(strFileOld)
    '            Catch ex As Exception

    '            End Try

    '            strFile = strFilePath & (("\winners_pb_" & drawNumberStr) & ".xml")

    '            Try
    '                reader2 = XmlReader.Create(strFile)
    '                Do While reader2.Read()
    '                    If (reader2.NodeType = XmlNodeType.Element AndAlso reader2.Name = "winners") Then
    '                        winnerStr = reader2.GetAttribute(3)
    '                    End If
    '                Loop
    '                reader2.Close()

    '            Catch ex As Exception

    '                errorCode = -1
    '                errorMessage = "The file PWRB WINNERS  -winners_pb_" & drawNumberPrevStr & ".xml- is not in " & strFilePath
    '                winnerStr = "0"

    '            End Try




    '        Case 13 'LIFE
    '            drawNumberStr = drawNumberStr.PadLeft(4, "0")
    '            drawNumberPrevStr = drawNumberPrevStr.PadLeft(4, "0")

    '            strFileOld = strFilePath & (("\winners_cl_" & drawNumberPrevStr) & ".xml")

    '            Try
    '                FileSystem.Kill(strFileOld)
    '            Catch ex As Exception

    '            End Try

    '            strFile = strFilePath & (("\winners_cl_" & drawNumberStr) & ".xml")

    '            Try
    '                reader2 = XmlReader.Create(strFile)
    '                Do While reader2.Read()
    '                    If (reader2.NodeType = XmlNodeType.Element AndAlso reader2.Name = "winners") Then
    '                        winnerStr = reader2.GetAttribute(3)
    '                        'winnerStr = ((reader2.GetAttribute(3) & " | ") & reader2.GetAttribute(4))

    '                    End If
    '                Loop
    '                reader2.Close()

    '            Catch ex As Exception

    '                errorCode = -1
    '                errorMessage = "The file LIFE WINNERS  -winners_cl_" & drawNumberPrevStr & ".xml- is not in " & strFilePath
    '                winnerStr = "0"

    '            End Try


    '    End Select
    '    Return winnerStr
    'End Function

    'Function IsDrawDay(ByVal codSysProduct As Integer, ByVal toKnowDate As DateTime) As Boolean
    '    Dim diaSorteoBool As Boolean
    '    Dim numDayDrawInt As Integer
    '    Dim dayNumberWeekInt As Integer
    '    Dim numrecordsInt As Integer
    '    Dim Parametros1(0) As MySqlParameter
    '    Dim row1 As DataRow
    '    Dim table1 As DataTable
    '    diaSorteoBool = False

    '    dayNumberWeekInt = Weekday(toKnowDate.AddDays(-1))
    '    Parametros1(0) = New MySqlParameter("@CodProductSistema", codSysProduct)
    '    table1 = dataSourceAccess.RunStructureProcedureReturnDtable("QDrawNumberXGame", Parametros1)
    '    numrecordsInt = table1.Rows.Count

    '    If (numrecordsInt > 0) Then
    '        For Each row1 In table1.Rows
    '            numDayDrawInt = row1("CodDayWeekDraw")
    '            If (dayNumberWeekInt = numDayDrawInt) Then
    '                diaSorteoBool = True
    '                Exit For
    '            End If
    '        Next
    '    End If

    '    IsDrawDay = diaSorteoBool

    'End Function

    'Function nextDayDraw(ByVal codSysProduct As Integer, ByVal toKnowDate As DateTime) As Date
    '    Dim encontro As Boolean = False
    '    Dim numDayDrawInt As Integer
    '    Dim dayNumberWeekInt As Integer
    '    Dim numrecordsInt As Integer
    '    Dim Parametros1(0) As MySqlParameter
    '    Dim row1 As DataRow
    '    Dim table1 As DataTable


    '    dayNumberWeekInt = Weekday(toKnowDate)
    '    Parametros1(0) = New MySqlParameter("@CodProductSistema", codSysProduct)
    '    table1 = dataSourceAccess.RunStructureProcedureReturnDtable("QDrawNumberXGame", Parametros1)
    '    numrecordsInt = table1.Rows.Count

    '    While Not encontro

    '        For Each row1 In table1.Rows
    '            numDayDrawInt = row1("CodDayWeekDraw")
    '            If (dayNumberWeekInt = numDayDrawInt) Then
    '                encontro = True
    '                Exit For
    '            End If
    '        Next
    '        If encontro Then
    '            nextDayDraw = toKnowDate

    '        Else
    '            toKnowDate = toKnowDate.AddDays(1)
    '            dayNumberWeekInt = Weekday(toKnowDate)
    '        End If

    '    End While


    'End Function

    'Function deleteHistoryDaily(ByVal idProductInt As Integer, ByVal cdcInt As Integer, ByRef msgErrorStr As String) As Integer
    '    Dim num1 As Integer
    '    Dim Parametros1(1) As MySqlParameter

    '    Parametros1(0) = New MySqlParameter("@CodGame", idProductInt)
    '    Parametros1(1) = New MySqlParameter("@cdc", cdcInt)

    '    num1 = dataSourceAccess.RunStoreProcedureParametersNonQuery("deleteHistoryDrawXproductXcdc", Parametros1, msgErrorStr)

    '    Return num1

    'End Function

    'Function updateOnCall(ByVal onCallName As String, ByVal levelInt As Integer, ByRef msgErrorStr As String) As Integer
    '    Dim num1 As Integer
    '    Dim Parametros1(1) As MySqlParameter

    '    Parametros1(0) = New MySqlParameter("@onCallNameStr", onCallName)
    '    Parametros1(1) = New MySqlParameter("@currentLevelInt", levelInt)

    '    num1 = dataSourceAccess.RunStoreProcedureParametersNonQuery("updateOnCall", Parametros1, msgErrorStr)

    '    Return num1

    'End Function



    'Sub saveHistoryDaily(ByVal newDrawInt As Integer, ByVal newCdcInt As Integer, ByVal productSysNameStr As String, ByVal idProductInt As Integer, ByVal winners As String, ByVal OldJackPotDbl As Double, ByVal dayWeekNameStr As String)
    '    Dim num1 As Integer
    '    Dim Parametros1(6) As MySqlParameter
    '    Dim msgErrorStr As String = ""

    '    Parametros1(0) = New MySqlParameter("@numDrawInt", newDrawInt)
    '    Parametros1(1) = New MySqlParameter("@numCDCInt", newCdcInt)
    '    Parametros1(2) = New MySqlParameter("@nomProductStr", productSysNameStr)
    '    Parametros1(3) = New MySqlParameter("@codigoProductInt", idProductInt)
    '    Parametros1(4) = New MySqlParameter("@numWinnerStr", winners)
    '    Parametros1(5) = New MySqlParameter("@JackpotDbl", OldJackPotDbl)
    '    Parametros1(6) = New MySqlParameter("@dayWeekNameStr", dayWeekNameStr)
    '    num1 = dataSourceAccess.RunStoreProcedureParametersNonQuery("InsertDraw", Parametros1, msgErrorStr)

    'End Sub

    'Sub saveLogLogin(ByVal userNameStr As String, ByVal updateDate As Date, ByVal descriptionActStr As String)
    '    Dim num1 As Integer
    '    Dim Parametros1(2) As MySqlParameter
    '    Dim msgErrorStr As String = ""

    '    Parametros1(0) = New MySqlParameter("@updateUser", userNameStr)
    '    Parametros1(1) = New MySqlParameter("@updateDate", updateDate)
    '    Parametros1(2) = New MySqlParameter("@descriptionAction", descriptionActStr)

    '    num1 = dataSourceAccess.RunStoreProcedureParametersNonQuery("insertLog_login", Parametros1, msgErrorStr)

    'End Sub



    'Sub updateJackPot(ByVal sysCodeProductInt As Integer, ByVal currentJackPotDbl As Double)
    '    Dim num1 As Integer
    '    Dim Parametros1(1) As MySqlParameter
    '    Dim msgErrorStr As String = ""

    '    Parametros1(0) = New MySqlParameter("@jackPotDbl", currentJackPotDbl)
    '    Parametros1(1) = New MySqlParameter("@sysCodeProductInt", sysCodeProductInt) 'System Code
    '    num1 = dataSourceAccess.RunStoreProcedureParametersNonQuery("UpdateJackPot", Parametros1, msgErrorStr)
    'End Sub

    'Sub updateAppVersion(ByVal appVersion As String, ByVal appName As String, ByVal userNameStr As String)
    '    Dim num1 As Integer
    '    Dim Parametros1(2) As MySqlParameter
    '    Dim msgErrorStr As String = ""

    '    Parametros1(0) = New MySqlParameter("@appversion", appVersion)
    '    Parametros1(1) = New MySqlParameter("@appname", appName)
    '    Parametros1(2) = New MySqlParameter("@userNameStr", usernameStr)
    '    num1 = dataSourceAccess.RunStoreProcedureParametersNonQuery("updateAppVersion", Parametros1, msgErrorStr)
    'End Sub




    'Function getLast2WinnersJackpotTbl(ByVal IdProductInt As Integer) As DataTable

    '    Dim Parametros1(0) As MySqlParameter
    '    Dim table1 As DataTable
    '    Parametros1(0) = New MySqlParameter("@codGame", IdProductInt)

    '    table1 = dataSourceAccess.RunStructureProcedureReturnDtable("Q2LastWinnersXGame", Parametros1)

    '    Return table1
    'End Function

    'Function getLastUserTbl(ByVal activityStr As String) As DataTable

    '    Dim Parametros1(0) As MySqlParameter
    '    Dim table1 As DataTable
    '    Parametros1(0) = New MySqlParameter("@activity", activityStr)

    '    table1 = dataSourceAccess.RunStructureProcedureReturnDtable("returnLastUser", Parametros1)

    '    Return table1
    'End Function

    'Function getInfoUserFromDatabase(ByVal userNameStr As String) As DataTable

    '    Dim Parametros1(0) As MySqlParameter
    '    Dim table1 As DataTable

    '    Parametros1(0) = New MySqlParameter("@username", userNameStr)
    '    table1 = dataSourceAccess.RunStructureProcedureReturnDtable("returnInfoUser", Parametros1)

    '    Return table1

    'End Function

    'Function getLastWinnersTake5Tbl(ByVal limitsInt As Integer) As DataTable

    '    Dim winnersTake5Tbl As New DataTable("TAKE5_WINNERS")
    '    Dim newRowTake5 As DataRow
    '    Dim Parametros1(0) As MySqlParameter
    '    Dim table1 As DataTable
    '    Dim contador As Integer

    '    Parametros1(0) = New MySqlParameter("@limite", limitsInt)

    '    winnersTake5Tbl = Utils.DespliegueTake5Winners

    '    table1 = dataSourceAccess.RunStructureProcedureReturnDtable("winnersTake5", Parametros1)
    '    For Each row1 In table1.Rows

    '        newRowTake5 = winnersTake5Tbl.NewRow()
    '        newRowTake5(0) = Trim(row1("dayWeekName"))
    '        newRowTake5(1) = row1("numCDC")
    '        'newRowTake5(2) = row1("numDraw")
    '        newRowTake5(2) = row1("Winners")
    '        winnersTake5Tbl.Rows.Add(newRowTake5)

    '    Next


    '    contador = limitsInt
    '    While contador < 7
    '        newRowTake5 = winnersTake5Tbl.NewRow()
    '        newRowTake5(0) = Trim(WeekdayName(contador + 1))
    '        winnersTake5Tbl.Rows.Add(newRowTake5)
    '        contador = contador + 1
    '    End While



    '    'contador = 7
    '    'While contador > limitsInt
    '    '    newRowTake5 = winnersTake5Tbl.NewRow()
    '    '    newRowTake5(0) = Trim(WeekdayName(contador))
    '    '    winnersTake5Tbl.Rows.Add(newRowTake5)
    '    '    contador = contador - 1
    '    'End While



    '    Return winnersTake5Tbl
    'End Function


    'Function readJackpotFromPaginaWeb(ByVal laUrl As String, ByVal sysCodeProductInt As Integer) As Double
    '    Dim jackpotDbl As Double = 0
    '    Dim num2 As Integer
    '    Dim num3 As Integer
    '    Dim reader1 As StreamReader
    '    Dim request1 As WebRequest
    '    Dim response1 As WebResponse
    '    Dim str2 As String = ""
    '    Dim str3 As String
    '    Dim str4 As String
    '    Dim resultStringFromWebPageStr As String


    '    request1 = WebRequest.Create(laUrl)
    '    response1 = request1.GetResponse()
    '    reader1 = New StreamReader(response1.GetResponseStream())
    '    resultStringFromWebPageStr = reader1.ReadToEnd()
    '    reader1.Close()
    '    response1.Close()

    '    Select Case sysCodeProductInt

    '        Case 8
    '            str2 = "ROS-NY Lotto"

    '        Case 12
    '            str2 = "ROS-Mega Millions"

    '        Case 15
    '            str2 = "ROS-Powerball"

    '        Case Else
    '            Return jackpotDbl

    '    End Select
    '    num2 = InStr(1, resultStringFromWebPageStr, str2, CompareMethod.Text)
    '    str3 = Mid(resultStringFromWebPageStr, (num2 - 400), 200)
    '    num3 = InStr(1, str3, "next-jackpot-amount'>$", CompareMethod.Text)
    '    str4 = Mid(str3, (num3 + 22), 14)
    '    num3 = InStr(1, str4, "<", CompareMethod.Text)
    '    str4 = Mid(str4, 1, (num3 - 1))
    '    jackpotDbl = CDbl(str4)

    '    Return jackpotDbl
    'End Function

    'Function getCurrentJackpotFromDatabase(ByVal IdProductInt As Integer) As Double
    '    Dim jackpotDbl As Double = 0
    '    Dim err1 As Integer

    '    Dim Parametros1(0) As MySqlParameter
    '    Dim msgErrorStr As String = ""

    '    Parametros1(0) = New MySqlParameter("@CodGame", IdProductInt)
    '    err1 = dataSourceAccess.RunStoreProcedureScalar("returnNewJackpot", Parametros1, jackpotDbl, msgErrorStr)
    '    Return jackpotDbl
    'End Function

    'Function readHostModeFile(ByVal strFileModeHost As String, ByRef hostNumberArray() As String, ByRef errMessageStr As String) As Boolean

    '    Dim reader1 As StreamReader
    '    Dim strLinea As String

    '    Try
    '        reader1 = New StreamReader(strFileModeHost, Encoding.UTF7)
    '        strLinea = ""
    '        strLinea = reader1.ReadLine()
    '        hostNumberArray = strLinea.Split("|")
    '        reader1.Close()
    '        Return True

    '    Catch ex As Exception
    '        errMessageStr = ex.Message
    '        Return False

    '    End Try

    'End Function


    'Function getSaveHostMode(ByVal strFileModeHost As String, ByRef hostNumberArray() As String, ByRef errMessageStr As String) As Integer

    '    Dim Parametros1(7) As MySqlParameter
    '    Dim num1 As Integer
    '    Dim resBoolHostConfi As Boolean

    '    If readHostModeFile(strFileModeHost, hostNumberArray, errMessageStr) Then

    '        resBoolHostConfi = validateLastHostConfi(hostNumberArray(0), hostNumberArray(1), hostNumberArray(2), hostNumberArray(3), hostNumberArray(4)) 'Return True if the last record has the same configuration of hosts

    '        If Not resBoolHostConfi Then
    '            Parametros1(0) = New MySqlParameter("@primaryHost", hostNumberArray(0))
    '            Parametros1(1) = New MySqlParameter("@secondaryHost", hostNumberArray(1))
    '            Parametros1(2) = New MySqlParameter("@spare1", hostNumberArray(2))
    '            Parametros1(3) = New MySqlParameter("@spare2", hostNumberArray(3))
    '            Parametros1(4) = New MySqlParameter("@spare3", hostNumberArray(4))
    '            Parametros1(5) = New MySqlParameter("@dateday", Now)
    '            Parametros1(6) = New MySqlParameter("@updateDate", Now)
    '            Parametros1(7) = New MySqlParameter("@updateUser", "System")
    '            num1 = dataSourceAccess.RunStoreProcedureParametersNonQuery("inserthostsystems", Parametros1, errMessageStr)
    '            If num1 = 0 Then
    '                Return 0

    '            Else
    '                Return 1
    '                'MsgBox("Error trying to insert Hosts: " & errMessageStr, MsgBoxStyle.Information, "Error")

    '            End If
    '        Else
    '            Return 2

    '        End If

    '    Else
    '        Return 3

    '    End If


    'End Function


    'Function validateEquals(ByVal priHostInt As Integer, ByVal secHostInt As Integer, ByVal spare1Int As Integer, ByVal spare2Int As Integer, ByVal spare3Int As Integer) As Boolean

    '    If (priHostInt = secHostInt) Or (priHostInt = spare1Int) Or (priHostInt = spare2Int) Or (priHostInt = spare3Int) Or (secHostInt = spare1Int) Or _
    '    (secHostInt = spare2Int) Or (secHostInt = spare3Int) Or (spare1Int = spare2Int) Or (spare1Int = spare3Int) Or (spare2Int = spare3Int) Then
    '        Return False
    '    Else
    '        Return True

    '    End If

    'End Function

    'Function validateLastHostConfi(ByVal priHostInt As Integer, ByVal secHostInt As Integer, ByVal spare1Int As Integer, ByVal spare2Int As Integer, ByVal spare3Int As Integer) As Boolean

    '    Dim tableHostConfi As DataTable
    '    Dim primaryHostDatabaseInt, secondaryHostDatabaseInt, spare1HostDatabaseInt, spare2HostDatabaseInt, spare3HostDatabaseInt As Integer

    '    tableHostConfi = dataSourceAccess.RunStoreQueryWithoutParameters("qhostsystems")

    '    primaryHostDatabaseInt = tableHostConfi.Rows(0).Item(0)
    '    secondaryHostDatabaseInt = tableHostConfi.Rows(0).Item(1)
    '    spare1HostDatabaseInt = tableHostConfi.Rows(0).Item(2)
    '    spare2HostDatabaseInt = tableHostConfi.Rows(0).Item(3)
    '    spare3HostDatabaseInt = tableHostConfi.Rows(0).Item(4)


    '    If (priHostInt = primaryHostDatabaseInt) And (secHostInt = secondaryHostDatabaseInt) And (spare1Int = spare1HostDatabaseInt) And (spare2Int = spare2HostDatabaseInt) And (spare3Int = spare3HostDatabaseInt) Then
    '        Return True
    '    Else
    '        Return False

    '    End If

    'End Function



    'Function returnColor(ByVal systemNumberInt As Integer) As Color

    '    Select Case systemNumberInt
    '        Case 1
    '            Return Color.Blue

    '        Case 2
    '            Return Color.Red

    '        Case 3
    '            Return Color.Green

    '        Case 4
    '            Return Color.Black

    '        Case 5
    '            Return Color.Purple

    '    End Select

    'End Function

    'Function UserAuthenticateActiveDirectory(ByVal Path As String, ByVal Username As String, ByVal Pass As String) As Boolean

    '    Dim dirSearch As DirectorySearcher
    '    Dim dirEntry As New DirectoryEntry(Path, Username, Pass, AuthenticationTypes.Secure)

    '    Try
    '        dirSearch = New DirectorySearcher(dirEntry)
    '        dirSearch.FindOne()
    '        Return True

    '    Catch ex As Exception

    '        Return False

    '    End Try


    'End Function


    ''Function getVersionApps() As DataTable

    ''    Dim versionAppsTbl As New DataTable("VERSION_APPS")
    ''    Dim newRowTake5 As DataRow
    ''    Dim Parametros1(0) As MySqlParameter
    ''    Dim table1 As DataTable
    ''    Dim contador As Integer
    ''    Parametros1(0) = New MySqlParameter("@limite", limitsInt)


    ''    winnersTake5Tbl = Utils.DespliegueTake5Winners
    ''    contador = 7
    ''    While contador > limitsInt
    ''        newRowTake5 = winnersTake5Tbl.NewRow()
    ''        newRowTake5(0) = Trim(WeekdayName(contador))
    ''        winnersTake5Tbl.Rows.Add(newRowTake5)
    ''        contador = contador - 1
    ''    End While

    ''    table1 = dataSourceAccess.RunStructureProcedureReturnDtable("winnersTake5", Parametros1)
    ''    For Each row1 In table1.Rows

    ''        newRowTake5 = winnersTake5Tbl.NewRow()
    ''        newRowTake5(0) = Trim(row1("dayWeekName"))
    ''        newRowTake5(1) = row1("numCDC")
    ''        newRowTake5(2) = row1("numDraw")
    ''        newRowTake5(3) = row1("Winners")
    ''        winnersTake5Tbl.Rows.Add(newRowTake5)

    ''    Next

    ''    Return winnersTake5Tbl
    ''End Function





End Module
