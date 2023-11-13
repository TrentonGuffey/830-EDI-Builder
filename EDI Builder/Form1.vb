Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Security.Cryptography
Public Class Form1
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Dim theFiles As New List(Of String)
        Dim files() As String = IO.Directory.GetFiles(Application.StartupPath)

        For Each file As String In files

            'MsgBox(Mid(file, file.Length - 3, 4))

            If Mid(file, file.Length - 3, 4) = ".edi" Then

                parseEDI(file)

            End If


        Next

        MsgBox("Finished")

    End Sub

    Private Sub parseEDI(ByVal theFileName As String)

        Dim outFileName As String = ""

        Try

            Dim locFile As String = ""
            Dim locNum As String = ""

            Dim theFile As String = ""
            Dim sbFile As New StringBuilder
            sbFile.AppendLine("PN" & "," & "Quantity" & "," & "Date_Wk" & "," & "Status" & "," & "Location" & "," & "PO")

            Using sr As New StreamReader(theFileName)
                theFile = sr.ReadToEnd
                'MsgBox(theFile)
            End Using




            Dim PN As String = ""
            Dim Quantity As String = ""
            Dim Date_Wk As String = ""
            Dim Status As String = ""
            Dim Location As String = ""
            Dim PO As String = ""

            For Each line As String In File.ReadLines(theFileName)

                'MsgBox(line)

                Dim lineArray() As String = line.Split(New String() {"*"}, StringSplitOptions.RemoveEmptyEntries)

                locFile = Mid(theFileName, InStrRev(theFileName, "\", theFileName.Count) + 1, 4).Replace(" ", "")
                If lineArray(1) = "ST" Then
                    Location = locFile & "-" & lineArray(lineArray.Count - 1)
                End If

                If lineArray(0) = "LIN" Then
                    'Part_Number = lineArray(2)
                    For j = 0 To lineArray.Count - 1
                        If lineArray(j) = "BP" Then
                            PN = lineArray(j + 1)
                        End If
                    Next
                End If

                If lineArray(0) = "LIN" Then
                    For h = 0 To lineArray.Count - 1
                        If lineArray(h) = "PO" Then
                            PO = lineArray(h + 1)
                        End If
                    Next
                End If

                If lineArray(0) = "FST" Then
                    Quantity = lineArray(1)
                    Status = lineArray(2).Replace("C", "Firm").Replace("D", "Planning")
                    Date_Wk = lineArray(4)
                    sbFile.AppendLine(PN & "," & Quantity & "," & Date_Wk & "," & Status & "," & Location & "," & PO)
                End If

            Next



            Dim finiFile As String = sbFile.ToString


            outFileName = theFileName
            outFileName = Mid(theFileName, 1, theFileName.Length - 4)
            outFileName = outFileName & ".csv"

            Using outfile As New StreamWriter(outFileName)
                outfile.Write(finiFile)
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



        Dim excelApp As New Excel.Application
        excelApp.DisplayAlerts = False
        Dim excelBook As Excel.Workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value)
        Dim theSheet As Excel.Worksheet = Nothing
        'try to open the workbook and a worksheet
        Try
            'MsgBox(outFileName)
            excelBook = excelApp.Workbooks.Open(outFileName)
            theSheet = CType(excelBook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)

            Dim sheet2 As Excel.Worksheet
            sheet2 = CType(excelBook.Worksheets.Add(), Excel.Worksheet)

            sheet2.Name = "ORDERING TABLE"

            Dim sheet3 As Excel.Worksheet
            sheet3 = CType(excelBook.Worksheets.Add(), Excel.Worksheet)

            sheet3.Name = "TOTALS"

            Dim sheet4 As Excel.Worksheet
            sheet4 = CType(excelBook.Worksheets.Add(), Excel.Worksheet)

            sheet4.Name = "M_CUTS"

            theSheet.Range("A1:F1").Font.Bold = True
            theSheet.Range("A1:F1").EntireColumn.AutoFit()

            theSheet.Range("A1:F1").Copy()
            sheet4.Paste()

            Dim range1 As Excel.Range
            range1 = theSheet.UsedRange

            With range1
                .AutoFilter(Field:=1, Criteria1:="=*M*")
                range1.Copy()
                sheet4.Paste()
                range1.EntireRow.Offset(1).Delete()
                If theSheet.AutoFilterMode = True Then
                    theSheet.AutoFilterMode = False
                End If
            End With

            Dim theValue As String = "value"
            Dim incR As Integer = 1
            Do Until theValue = ""
                theValue = theSheet.Range("A" & incR).Value
                incR += 1
            Loop
            incR = incR - 2

            Dim lastRow As Long = theSheet.Cells.Rows.Count
            Dim lastCol As Long = theSheet.Cells.Columns.Count

            Dim dataRange As Excel.Range = theSheet.Range("A1:F" & incR)
            Dim ptName As String = "MyPivotTable"
            sheet2.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, dataRange, sheet2.Cells(3, 1), ptName)
            sheet2.Select()
            sheet3.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, dataRange, sheet3.Cells(3, 1), ptName)
            sheet3.Select()

            'the pivot table version has a lot to do with the formatting - search Excel.XlPivotTableVersionList.xlPivotTableVersion14
            Dim pt As Excel.PivotTable = sheet2.PivotTables(1)
            With pt
                .TableStyle = Excel.XlPivotTableVersionList.xlPivotTableVersion14
                .TableStyle2 = "PivotStyleLight16"
                .InGridDropZones = False
            End With

            With pt.PivotFields("Date_Wk")
                .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                .Position = 1
            End With
            With pt.PivotFields("Location")
                .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                .Position = 1
            End With
            With pt.PivotFields("PN")
                .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                .Position = 2
            End With

            With pt.PivotFields("PO")
                .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                .Position = 3
            End With

            With pt.PivotFields("Quantity")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Position = 1
            End With
            For Each pField In pt.PivotFields
                pField.Subtotals(1) = True
                pField.Subtotals(1) = False
            Next pField
            sheet2.Columns("A:C").EntireColumn.AutoFit()
            Dim pt2 As Excel.PivotTable = sheet3.PivotTables(1)
            With pt2
                .TableStyle = Excel.XlPivotTableVersionList.xlPivotTableVersion14
                .TableStyle2 = "PivotStyleLight16"
                .InGridDropZones = False
            End With

            With pt2.PivotFields("Date_Wk")
                .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                .Position = 1
            End With
            With pt2.PivotFields("PN")
                .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                .Position = 1
            End With

            With pt2.PivotFields("Quantity")
                .Orientation = Excel.XlPivotFieldOrientation.xlDataField
                .Position = 1
            End With
            theSheet.PageSetup.LeftHeader = "&F"
            sheet2.PageSetup.LeftHeader = "&F"
            sheet3.PageSetup.LeftHeader = "&F"
            sheet4.PageSetup.LeftHeader = "&F"

            theSheet.PageSetup.LeftMargin = "18"
            theSheet.PageSetup.RightMargin = "18"
            theSheet.PageSetup.TopMargin = "54"
            theSheet.PageSetup.BottomMargin = "54"
            theSheet.PageSetup.HeaderMargin = "21.5"
            theSheet.PageSetup.FooterMargin = "21.5"
            sheet2.PageSetup.LeftMargin = "18"
            sheet2.PageSetup.RightMargin = "18"
            sheet2.PageSetup.TopMargin = "54"
            sheet2.PageSetup.BottomMargin = "54"
            sheet2.PageSetup.HeaderMargin = "21.5"
            sheet2.PageSetup.FooterMargin = "21.5"
            sheet3.PageSetup.LeftMargin = "18"
            sheet3.PageSetup.RightMargin = "18"
            sheet3.PageSetup.TopMargin = "54"
            sheet3.PageSetup.BottomMargin = "54"
            sheet3.PageSetup.HeaderMargin = "21.5"
            sheet3.PageSetup.FooterMargin = "21.5"
            sheet4.PageSetup.LeftMargin = "18"
            sheet4.PageSetup.RightMargin = "18"
            sheet4.PageSetup.TopMargin = "54"
            sheet4.PageSetup.BottomMargin = "54"
            sheet4.PageSetup.HeaderMargin = "21.5"
            sheet4.PageSetup.FooterMargin = "21.5"

            Dim range2 As Excel.Range
            range2 = sheet2.UsedRange
            range2.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            range2.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Color = 2
            range2.Borders(Excel.XlBordersIndex.xlInsideVertical).Color = 2

            Dim range3 As Excel.Range
            range3 = sheet3.UsedRange
            range3.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            range3.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Color = 2
            range3.Borders(Excel.XlBordersIndex.xlInsideVertical).Color = 2

            Dim range4 As Excel.Range
            range4 = sheet4.UsedRange
            range4.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            range4.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Color = 2
            range4.Borders(Excel.XlBordersIndex.xlInsideVertical).Color = 2

            sheet2.Range("A3").Value = "."
            sheet3.Range("A3").Value = "."
            sheet2.Range("A4").End(Excel.XlDirection.xlToRight).Value = "TOTALS"
            sheet3.Range("A4").End(Excel.XlDirection.xlToRight).Value = "TOTALS"

            theSheet.Columns("F:F").NumberFormat = "0"
            sheet2.Columns("C:C").NumberFormat = "0"

            theSheet.Columns.AutoFit()
            sheet2.Columns.AutoFit()
            sheet3.Columns.AutoFit()
            sheet4.Columns.AutoFit()

            sheet2.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
            sheet3.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

            theSheet.Range("A:B").HorizontalAlignment = Excel.Constants.xlCenter
            sheet2.Range("B:C").HorizontalAlignment = Excel.Constants.xlCenter
            sheet4.Range("A:B").HorizontalAlignment = Excel.Constants.xlCenter


        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            excelBook.SaveAs(outFileName.Replace("csv", "xlsx"), FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
            'MAKE SURE TO KILL ALL INSTANCES BEFORE QUITING! if you fail to do this
            'The service (excel.exe) will continue to run
            NAR(theSheet)
            excelBook.Close(False)
            NAR(excelBook)
            excelApp.Workbooks.Close()
            NAR(excelApp.Workbooks)
            'quit and dispose app
            excelApp.Quit()
            NAR(excelApp)
            'VERY IMPORTANT
            GC.Collect()
            My.Computer.FileSystem.DeleteFile(outFileName)
        End Try
    End Sub

    Private Sub NAR(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try

    End Sub

End Class
