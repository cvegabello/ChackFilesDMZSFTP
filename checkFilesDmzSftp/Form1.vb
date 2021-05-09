
Imports System.Text
Imports System.IO
'Imports Microsoft.Office.Interop.Excel


Public Class Form1

    

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dateSelected As Date
        Dim cdcInt As Integer
        Dim hourInt, IsExtraFileInt As Integer
        'Dim tableRes As DataTable
        'Dim fecha As Date
        ComboBox1.Text = Utils.GetSetting(appName, "Server", "").ToString()
        nameFileTxt.Text = Utils.GetSetting(appName, "FileName", "").ToString()


        IsExtraFileInt = InStr(nameFileTxt.Text, "listFilesExistExtraOps")
        hourInt = Now.TimeOfDay.TotalHours
        If hourInt <= 12 And IsExtraFileInt <> 0 Then
            DateTimePicker1.Value = Date.Today.AddDays(-1)
        Else
            DateTimePicker1.Value = Now

        End If

        ' dateTxt.Text = Format(Now, "MM/dd/yyyy")
        'tableRes = Utils.getInfoFileTbl(fecha)
        'grdCheckFiles.DataSource = tableRes

        dateSelected = DateTimePicker1.Value
        cdcInt = returnCDC(dateSelected)
        cdcLbl.Text = cdcInt
        Button1.Enabled = False
        formatGrid()
    End Sub


    Private Sub formatGrid()

        Dim DataGridViewCellStyle1 As DataGridViewCellStyle = New DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As DataGridViewCellStyle = New DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As DataGridViewCellStyle = New DataGridViewCellStyle
        'Dim DataGridViewCellStyle4 As DataGridViewCellStyle = New DataGridViewCellStyle
        Dim Column1 As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn
        Dim Column2 As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn
        Dim Column3 As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn
        Dim Column4 As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn
        Dim Column5 As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn


        Column1 = grdCheckFiles.Columns(0)
        Column2 = grdCheckFiles.Columns(1)
        Column3 = grdCheckFiles.Columns(2)
        Column4 = grdCheckFiles.Columns(3)
        Column5 = grdCheckFiles.Columns(4)



        'File Name
        '
        DataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft
        'DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle1.BackColor = Color.DarkKhaki
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = Color.DarkKhaki
        DataGridViewCellStyle1.SelectionForeColor = Drawing.Color.Black

        Column1.DefaultCellStyle = DataGridViewCellStyle1
        Column1.HeaderText = "FILE NAME"

        Column1.Resizable = DataGridViewTriState.[True]
        Column1.SortMode = DataGridViewColumnSortMode.NotSortable
        Column1.Width = 440



        'Location
        ''
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.BackColor = Color.DarkKhaki
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = Drawing.Color.Black
        Column2.DefaultCellStyle = DataGridViewCellStyle2
        Column2.HeaderText = "LOCATION"

        Column2.Resizable = DataGridViewTriState.[True]
        Column2.SortMode = DataGridViewColumnSortMode.NotSortable
        Column2.Width = 485


        'Size
        '
        'DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        'DataGridViewCellStyle2.BackColor = Color.DarkKhaki
        'DataGridViewCellStyle2.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'DataGridViewCellStyle2.ForeColor = Drawing.Color.Maroon
        'DataGridViewCellStyle2.Format = "#,##0"
        Column3.DefaultCellStyle = DataGridViewCellStyle2
        Column3.HeaderText = "FILE SIZE"
        Column3.Resizable = DataGridViewTriState.[True]
        Column3.SortMode = DataGridViewColumnSortMode.NotSortable
        Column3.Width = 100


        Column4.DefaultCellStyle = DataGridViewCellStyle2
        Column4.HeaderText = "LAST MODIFIED"
        Column4.Resizable = DataGridViewTriState.[True]
        Column4.SortMode = DataGridViewColumnSortMode.Automatic
        Column4.Width = 250



        ''Exist Yes or No
        ''
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = Color.DarkKhaki
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = Drawing.Color.Black
        Column5.DefaultCellStyle = DataGridViewCellStyle3
        Column5.HeaderText = "EXIST"

        Column5.Resizable = DataGridViewTriState.[True]
        Column5.SortMode = DataGridViewColumnSortMode.NotSortable
        Column5.Width = 80


        ''colTime
        ''
        'DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        'DataGridViewCellStyle4.BackColor = Color.DarkKhaki
        ''DataGridViewCellStyle4.Font = New System.Drawing.Font("Calibri", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'DataGridViewCellStyle4.ForeColor = Drawing.Color.Maroon
        'Column4.DefaultCellStyle = DataGridViewCellStyle4
        'Column4.HeaderText = ""

        'Column4.Resizable = DataGridViewTriState.[True]
        'Column4.SortMode = DataGridViewColumnSortMode.NotSortable
        'Column4.Width = 170




    End Sub






    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        If (result = DialogResult.OK) Then
            nameFileTxt.Text = Trim(OpenFileDialog1.FileName)

        End If
        okBtn.Focus()
    End Sub

    Private Sub ComboBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComboBox1.MouseClick
        'MsgBox("Hello")
        ComboBox1.DroppedDown = True
    End Sub

    
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim dateSelected As Date
        Dim cdcInt As Integer

        dateSelected = DateTimePicker1.Value
        cdcInt = returnCDC(dateSelected)
        cdcLbl.Text = cdcInt
    End Sub

    Private Sub cancelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelBtn.Click
        Me.Dispose()
    End Sub

    Private Sub okBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okBtn.Click
        Dim sftp As New Chilkat.SFtp
        Dim errMessageStr As String = ""
        Dim errCod, contadorInt, numRowsInt, IsExtraFileInt, hourInt As Integer
        Dim rowStr As String()
        Dim sizeStr, serverStr, cdcStr As String
        Dim fecha As Date
        Dim flagExtra As Boolean

        Dim dateCrFileDate As DateTime

        Dim reader1 As StreamReader
        Dim strLine, strValue, fileNameStr, cdcOTGStr As String
        Dim monthStr, dayStr, yearStr, dateFileStr As String

        Dim substrings() As String
        Dim DataGridViewCellStyle1 As DataGridViewCellStyle = New DataGridViewCellStyle



        IsExtraFileInt = InStr(nameFileTxt.Text, "listFilesExistExtraOps")
        hourInt = Now.TimeOfDay.TotalHours
        If hourInt <= 12 And IsExtraFileInt <> 0 Then
            DateTimePicker1.Value = Date.Today.AddDays(-1)
        Else
            DateTimePicker1.Value = Now

        End If



        'IsExtraFileInt = InStr(nameFileTxt.Text, "listFilesExistExtraOps")
        'hourInt = Now.TimeOfDay.TotalHours
        'If hourInt <= 12 And IsExtraFileInt <> 0 Then
        '    If DateTimePicker1.Value <> Date.Today.AddDays(-1) Then
        '        MsgBox("The date selected is NOT the current date.")
        '    End If
        'Else
        '    If DateTimePicker1.Value <> Now Then
        '        MsgBox("The date selected is NOT the current date.")
        '    End If

        'End If

        grdCheckFiles.Cursor = Cursors.WaitCursor
        grdCheckFiles.Rows.Clear()

        DataGridViewCellStyle1.BackColor = Color.Yellow
        DataGridViewCellStyle1.ForeColor = Color.Red
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

        'cdcStr = cdcLbl.Text.PadLeft(5, "0")


        strValue = ""
        serverStr = Trim(ComboBox1.Text)

        If nameFileTxt.Text.Contains("listFilesExistExtraOps") Then
            flagExtra = True
        Else
            flagExtra = False
        End If

        Select Case serverStr

            Case ""
                MsgBox("You must select a server", MsgBoxStyle.OkOnly, "Select a server")
                Exit Sub


            Case "CAT_SFTP (10.162.33.15)"
                errCod = connectSFTP(sftp, "10.162.33.15", "22", "prosys", "prosys", errMessageStr)


            Case "DMZSFTP1 (10.1.6.184)"
                If flagExtra Then
                    errCod = connectSFTP(sftp, "10.1.6.184", "22", "glxfer", "NYL3dger2018!", errMessageStr)
                Else
                    errCod = connectSFTP(sftp, "10.1.6.184", "22", "xfer", "Welcome1", errMessageStr)
                End If

            Case "NYSFTP1 (10.1.5.182)"

                errCod = connectSFTP(sftp, "10.1.5.182", "22", "xfer", "Welcome1", errMessageStr)

            Case "DMZSFTP2 (10.2.6.185)"
                If flagExtra Then
                    errCod = connectSFTP(sftp, "10.2.6.185", "22", "glxfer", "NYL3dger2018!", errMessageStr)
                Else
                    errCod = connectSFTP(sftp, "10.2.6.185", "22", "xfer", "Welcome1", errMessageStr)
                End If


        End Select




        Select Case errCod

            Case 1
                MsgBox("Error 1")

            Case 2
                MsgBox("Error 2")

            Case 3

                MsgBox("Error 3")

            Case 4
                MsgBox("Error 4")

            Case 0

                If nameFileTxt.Text <> "" Then
                    Try
                        reader1 = New StreamReader(nameFileTxt.Text, Encoding.UTF7)
                        'reader1 = New StreamReader("listFilesExistTest.csv", Encoding.UTF7)
                        strLine = ""
                        contadorInt = 0


                        Do While Not (strLine Is Nothing)
                            strLine = reader1.ReadLine()

                            If Not (strLine Is Nothing) Then
                                substrings = strLine.Split(",")

                                If (substrings(0).Contains("*")) Then 'CDC IGT
                                    cdcStr = cdcLbl.Text.PadLeft(5, "0")
                                    fileNameStr = substrings(0).Replace("*", cdcStr)
                                ElseIf (substrings(0).Contains("%")) Then 'CDC IGT + 1
                                    cdcStr = CStr(CInt(cdcLbl.Text) + 1)
                                    cdcStr = cdcStr.PadLeft(5, "0")
                                    fileNameStr = substrings(0).Replace("%", cdcStr)
                                ElseIf (substrings(0).Contains("<")) Then 'CDC IGT - 1
                                    cdcStr = CStr(CInt(cdcLbl.Text) - 1)
                                    cdcStr = cdcStr.PadLeft(5, "0")
                                    fileNameStr = substrings(0).Replace("<", cdcStr)



                                ElseIf (substrings(0).Contains("$")) Then 'Current date format yymmdd

                                    fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
                                    monthStr = CStr(fecha.Month).PadLeft(2, "0")
                                    dayStr = CStr(fecha.Day).PadLeft(2, "0")
                                    yearStr = Mid(CStr(fecha.Year), 3).PadLeft(2, "0")
                                    dateFileStr = yearStr + monthStr + dayStr
                                    fileNameStr = substrings(0).Replace("$", dateFileStr)

                                ElseIf (substrings(0).Contains("(")) Then 'Current date format yyyymmdd

                                    fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
                                    monthStr = CStr(fecha.Month).PadLeft(2, "0")
                                    dayStr = CStr(fecha.Day).PadLeft(2, "0")
                                    yearStr = CStr(fecha.Year)
                                    dateFileStr = yearStr + monthStr + dayStr
                                    fileNameStr = substrings(0).Replace("(", dateFileStr)

                                ElseIf (substrings(0).Contains("-")) Then 'Yesterday date format yymmdd
                                    fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
                                    fecha = fecha.AddDays(-1)
                                    monthStr = CStr(fecha.Month).PadLeft(2, "0")
                                    dayStr = CStr(fecha.Day).PadLeft(2, "0")
                                    yearStr = Mid(CStr(fecha.Year), 3).PadLeft(2, "0")
                                    dateFileStr = yearStr + monthStr + dayStr
                                    fileNameStr = substrings(0).Replace("-", dateFileStr)

                                ElseIf (substrings(0).Contains("^")) Then 'Tomorrow
                                    fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
                                    fecha = fecha.AddDays(1)
                                    monthStr = CStr(fecha.Month).PadLeft(2, "0")
                                    dayStr = CStr(fecha.Day).PadLeft(2, "0")
                                    yearStr = Mid(CStr(fecha.Year), 3).PadLeft(2, "0")
                                    dateFileStr = yearStr + monthStr + dayStr
                                    fileNameStr = substrings(0).Replace("^", dateFileStr)

                                ElseIf (substrings(0).Contains("@")) Then 'CDC OTG
                                    fecha = DateTimePicker1.Value
                                    cdcOTGStr = CStr(returnCDC_OTG(fecha)).PadLeft(5, "0")
                                    fileNameStr = substrings(0).Replace("@", cdcOTGStr)

                                ElseIf (substrings(0).Contains("?")) Then 'Previous CDC OTG
                                    fecha = DateTimePicker1.Value
                                    cdcOTGStr = CStr(returnCDC_OTG(fecha) - 1).PadLeft(5, "0")
                                    fileNameStr = substrings(0).Replace("?", cdcOTGStr)


                                ElseIf (substrings(0).Contains("#")) Then 'Date yyyymmdd OTG
                                    fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
                                    monthStr = CStr(fecha.Month).PadLeft(2, "0")
                                    dayStr = CStr(fecha.Day).PadLeft(2, "0")
                                    yearStr = CStr(fecha.Year)
                                    dateFileStr = yearStr + monthStr + dayStr
                                    fileNameStr = substrings(0).Replace("#", dateFileStr)

                                ElseIf (substrings(0).Contains("!")) Then 'Date yesterday con formato yyyymmdd OTG
                                    fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
                                    fecha = fecha.AddDays(-1)
                                    monthStr = CStr(fecha.Month).PadLeft(2, "0")
                                    dayStr = CStr(fecha.Day).PadLeft(2, "0")
                                    yearStr = CStr(fecha.Year)
                                    dateFileStr = yearStr + monthStr + dayStr
                                    fileNameStr = substrings(0).Replace("!", dateFileStr)

                                Else
                                    fileNameStr = substrings(0)

                                End If


                                If (sftp.FileExists(substrings(1) & "/" & fileNameStr, True) = 1) Then
                                    sizeStr = sftp.GetFileSizeStr(substrings(1) & "/" & fileNameStr, True, False)
                                    'sizeStr = FormatNumber(sizeStr, "0")
                                    dateCrFileDate = sftp.GetFileLastModifiedStr(substrings(1) & "/" & fileNameStr, True, False)
                                    rowStr = New String() {fileNameStr, substrings(1), sizeStr, dateCrFileDate, "YES"}
                                    grdCheckFiles.Rows.Add(rowStr)
                                    contadorInt += 1


                                Else
                                    rowStr = New String() {fileNameStr, substrings(1), "NA", "NA", "NO"}
                                    grdCheckFiles.Rows.Add(rowStr)
                                    contadorInt += 1
                                    grdCheckFiles.Rows(contadorInt - 1).DefaultCellStyle = DataGridViewCellStyle1

                                End If


                            End If
                        Loop
                        reader1.Close()
                        grdCheckFiles.ClearSelection()

                    Catch ex As Exception
                        MsgBox("Error reading file. " & ex.Message, MsgBoxStyle.Information, "Error")

                    End Try
                    'grdCheckFiles.Rows(2).DefaultCellStyle = DataGridViewCellStyle1

                Else
                    MsgBox("You must select a file.", MsgBoxStyle.OkOnly, "Select a file.")
                End If

        End Select
        sftp.Disconnect()

        Utils.SalvarSetting(appName, "Server", ComboBox1.Text)
        Utils.SalvarSetting(appName, "FileName", nameFileTxt.Text)
        grdCheckFiles.Cursor = Cursors.Default

        numRowsInt = grdCheckFiles.RowCount()

        If numRowsInt > 0 Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False

        End If

    End Sub

    Private Sub ComboBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComboBox1.MouseMove
        Cursor = Cursors.Arrow
    End Sub

    Private Sub ComboBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComboBox1.MouseUp
        Cursor = Cursors.Arrow
    End Sub


    Private Sub exportCsvBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exportCsvBtn.Click
        Dim numRowsInt, i, resCode As Integer
        Dim nameFileStr, emailStr, emailInputStr As String
        Dim file As System.IO.StreamWriter
        Dim monthStr, dayStr, yearStr, dateFileStr As String
        Dim fecha As Date

        fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year)
        dateFileStr = yearStr + monthStr + dayStr
        If nameFileTxt.Text.Contains("listOTGFilesAM") Then
            SaveFileDialog1.FileName = "opc" & dateFileStr & "AM.csv"
        End If


        numRowsInt = grdCheckFiles.RowCount()

        Dim result As DialogResult = SaveFileDialog1.ShowDialog()

        If (result = DialogResult.OK) Then

            nameFileStr = Trim(SaveFileDialog1.FileName)
            file = My.Computer.FileSystem.OpenTextFileWriter(nameFileStr, False, ASCIIEncoding.ASCII)
            'file.WriteLine("REPORT,DIRECTORY,FILE SIZE,TIME TRANSFERRED,EXIST,SERVER")

            For i = 0 To numRowsInt - 1
                'file.WriteLine(CStr(grdCheckFiles.Rows(i).Cells(0).Value) & "," & CStr(grdCheckFiles.Rows(i).Cells(1).Value) & "," & CStr(grdCheckFiles.Rows(i).Cells(2).Value) & "," & CStr(grdCheckFiles.Rows(i).Cells(3).Value) & "," & CStr(grdCheckFiles.Rows(i).Cells(4).Value) & ", " & ComboBox1.Text)
                file.WriteLine(CStr(grdCheckFiles.Rows(i).Cells(0).Value) & "," & CStr(grdCheckFiles.Rows(i).Cells(3).Value) & "," & CStr(grdCheckFiles.Rows(i).Cells(2).Value))
            Next

            file.Close()

            MsgBox("The information has been saved successfully in '" & nameFileStr & "'." & vbCr & vbCr & "Do you want this file to be sent to you email?", MsgBoxStyle.YesNo, "Export to CSV successful.")

            If MsgBoxResult.Yes Then
                emailInputStr = Utils.GetSettingEmail(appName, "email", "").ToString
                emailStr = InputBox("Please enter your email.", "Enter your email", emailInputStr)
                Utils.SalvarSettingEmail(appName, "email", emailStr)
                If emailStr <> "" Then
                    Cursor = Cursors.WaitCursor
                    resCode = sent_Email("156.24.14.132", emailStr, nameFileStr, "Check Files OTG " & fecha, "Please see the file attached.")
                    Cursor = Cursors.Default
                    If resCode = 0 Then
                        MessageBox.Show("The file was sent to your email.", "File sent to eMail", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Error!", "Error Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End If

        End If


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim xlibro As Microsoft.Office.Interop.Excel.Application
        Dim strPathFile, fileNameStr, monthStr, dayStr, yearStr, initialStr, emailInputStr, emailStr, fileNameTemplateStr, pathFileNameStr As String
        Dim numRowsInt, resCode As Integer
        Dim fecha As Date


        If nameFileTxt.Text.Contains("listOnePlaceFiles") Then
            initialStr = InputBox("Please enter your initials.", "Enter your initials", "")

            If initialStr <> "" Then
                Cursor = Cursors.WaitCursor
                fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
                strPathFile = System.Windows.Forms.Application.StartupPath
                fileNameTemplateStr = strPathFile & "\OnePlaceFile_Transfers_Checklist_TEMPLATE.xlsx"

                xlibro = CreateObject("Excel.Application")
                xlibro.Workbooks.Open(fileNameTemplateStr)
                xlibro.Visible = True

                xlibro.Sheets("OP Files").Select()
                xlibro.Cells(1, 10) = fecha

                numRowsInt = grdCheckFiles.RowCount()

                For i = 0 To numRowsInt - 2
                    xlibro.Cells(i + 2, 2) = CStr(grdCheckFiles.Rows(i).Cells(0).Value)
                    xlibro.Cells(i + 2, 4) = CStr(grdCheckFiles.Rows(i).Cells(1).Value)
                    xlibro.Cells(i + 2, 5) = CStr(grdCheckFiles.Rows(i).Cells(3).Value)
                    xlibro.Cells(i + 2, 6) = CStr(grdCheckFiles.Rows(i).Cells(2).Value)
                    xlibro.Cells(i + 2, 7) = initialStr
                Next

                fecha = Format(Now, "MM/dd/yyyy HH:mm:ss")
                monthStr = CStr(fecha.Month).PadLeft(2, "0")
                dayStr = CStr(fecha.Day).PadLeft(2, "0")
                yearStr = CStr(fecha.Year).PadLeft(4, "0")
                fileNameStr = "OnePlaceFile_Transfers_Checklist_" & yearStr & monthStr & dayStr & ".xlsx"
                pathFileNameStr = strPathFile & "\" & "OutputExcelFiles" & "\" & fileNameStr
                xlibro.ActiveWorkbook.SaveAs(pathFileNameStr)
                xlibro.Workbooks.Close()
                xlibro.Quit()
                Cursor = Cursors.Default
                emailInputStr = Utils.GetSettingEmailOnePlace(appName, "emailOnePlace", "").ToString
                emailStr = InputBox("Please enter recipients email, separated by commas.", "Enter recipients email", emailInputStr)
                If emailStr <> "" Then
                    Utils.SalvarSettingEmailOnePlace(appName, "emailOnePlace", emailStr)
                    Cursor = Cursors.WaitCursor
                    resCode = sent_Email("156.24.14.132", emailStr, pathFileNameStr, "OnePlace files transmitted " & fecha, "Please see the file attached." & vbCr & vbCr & "OPS, please save this to \\NYCR\nygroups\OPS\OnePlace")
                    Cursor = Cursors.Default
                    If resCode = 0 Then
                        MessageBox.Show("The file was sent to the recipients.", "File sent to eMail", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Error!", "Error Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End If

        Else
            MsgBox("The information is not OnePlace Files", MsgBoxStyle.Information, "CheckFilesDMZSftp")
        End If



    End Sub
End Class
