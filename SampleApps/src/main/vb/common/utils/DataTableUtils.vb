Imports Excel = Microsoft.Office.Interop.Excel

Public Module DataTableUtils
    '（注意）Excelシートの行、列のインデックスは１から始まる
    Private Const ROW_OFFSET As Integer = 2 'エクセルシートにデータを書き出す行番号
    '（注意）Excelシートの行、列のインデックスは１から始まる
    Private Const COLUMN_OFFSET As Integer = 1 'エクセルシートにデータを書き出す列番号

    '''<summary>
    '''ハッシュテーブルをエクセルファイルから取得します
    ''' Hashtable：key=SheetName,value=DataTable
    '''</summary>
    Public Function LoadDataSetFromExcel(ByVal filePath As String, Optional ByRef ds As DataSet = Nothing) As DataSet

        If IsNothing(ds) Then
            ds = New DataSet
        End If


        Dim xlsApplication As New Excel.Application
        Dim xlsBook As Excel.Workbook = Nothing
        Dim xlsSheets As Excel.Sheets = Nothing
        Dim xlsSheet As Excel.Worksheet = Nothing
        Dim xlsRange As Excel.Range = Nothing

        xlsApplication.DisplayAlerts = False   '保存時の確認ダイアログを表示しない
        xlsBook = xlsApplication.Workbooks.Open(filePath)
        Dim sheetList As New List(Of String)


        Try
            For Each sheet As Excel.Worksheet In xlsBook.Sheets
                sheetList.Add(sheet.Name)
            Next

            xlsBook.Close(False)
            xlsApplication.Quit()
        Catch ex As Exception
            Console.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            'エクセル関係のオブジェクトは必ず解放すること

            ReleaseComObject(DirectCast(xlsRange, Object))
            ReleaseComObject(DirectCast(xlsSheet, Object))
            ReleaseComObject(DirectCast(xlsSheets, Object))
            ReleaseComObject(DirectCast(xlsBook, Object))
            ReleaseComObject(DirectCast(xlsApplication, Object))
        End Try


        For Each sheetName As String In sheetList

            Dim dt As DataTable = getDataTableFromExcelSheet(filePath, sheetName)
            dt.TableName = sheetName
            ds.Tables.Add(dt)
        Next


        Return ds
    End Function
    '''<summary>
    '''ハッシュテーブルをエクセルファイルから取得します
    ''' Hashtable：key=SheetName,value=DataTable
    '''</summary>
    Public Function LoadHashTableFromExcel(ByVal filePath As String, Optional ByRef ht As Dictionary(Of String, DataTable) = Nothing) As Dictionary(Of String, DataTable)

        If IsNothing(ht) Then
            ht = New Dictionary(Of String, DataTable)
        End If


        Dim xlsApplication As New Excel.Application
        Dim xlsBook As Excel.Workbook = Nothing
        Dim xlsSheets As Excel.Sheets = Nothing
        Dim xlsSheet As Excel.Worksheet = Nothing
        Dim xlsRange As Excel.Range = Nothing

        xlsApplication.DisplayAlerts = False   '保存時の確認ダイアログを表示しない
        xlsBook = xlsApplication.Workbooks.Open(filePath)
        Dim sheetList As New List(Of String)


        Try
            For Each sheet As Excel.Worksheet In xlsBook.Sheets
                sheetList.Add(sheet.Name)
            Next

            xlsBook.Close(False)
            xlsApplication.Quit()
        Catch ex As Exception
            Console.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            'エクセル関係のオブジェクトは必ず解放すること

            ReleaseComObject(DirectCast(xlsRange, Object))
            ReleaseComObject(DirectCast(xlsSheet, Object))
            ReleaseComObject(DirectCast(xlsSheets, Object))
            ReleaseComObject(DirectCast(xlsBook, Object))
            ReleaseComObject(DirectCast(xlsApplication, Object))
        End Try


        For Each sheetName As String In sheetList
            Dim dt As DataTable = getDataTableFromExcelSheet(filePath, sheetName)
            ht.Add(sheetName, dt)
        Next


        Return ht
    End Function


    Private Function getDataTableFromExcelSheet(ByVal filePath As String, ByVal sheetName As String) As DataTable
        Dim Con As New OleDb.OleDbConnection
        Dim Command As New OleDb.OleDbCommand()
        Dim oDataTable1 As DataTable = New DataTable
        Dim ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0; " &
            "Data Source=" & filePath & ";" & "Extended Properties=""Excel 12.0;HDR=YES;"""
        '条件を指定してデータを取得したい場合
        Dim where As String = ""

        Try
            Dim oDataAdapter As New OleDb.OleDbDataAdapter
            Con.ConnectionString = ConnectionString
            Command.Connection = Con
            Command.CommandText = "SELECT * FROM [" & sheetName & "$]" & where
            oDataAdapter.SelectCommand = Command
            oDataAdapter.Fill(oDataTable1)

        Catch ex As Exception
            'エラー処理
            Throw
        Finally
            If Not Command Is Nothing Then
                Command.Dispose()
            End If
            If Not Con Is Nothing Then
                Con.Close()
                Con.Dispose()
            End If
        End Try

        Return oDataTable1

    End Function

    '''<summary>
    '''DataSetをエクセルファイルに出力します
    '''</summary>
    Public Function CreateExcelFromDataSet(ByVal ds As DataSet, ByVal filePath As String) As Boolean
        If IO.File.Exists(filePath) Then
            IO.File.Delete(filePath)
        End If


        Dim xlsApplication As New Excel.Application
        Dim xlsBook As Excel.Workbook = Nothing
        Dim xlsSheets As Excel.Sheets = Nothing
        Dim xlsSheet As Excel.Worksheet = Nothing
        Dim xlsRange As Excel.Range = Nothing

        xlsApplication.DisplayAlerts = False   '保存時の確認ダイアログを表示しない
        xlsBook = xlsApplication.Workbooks.Add()
        Dim tableNameList As New List(Of String)


        Try
            For Each table As DataTable In ds.Tables
                tableNameList.Add(table.TableName)
            Next
            tableNameList.Sort()
            tableNameList.Reverse()


            Dim dt As DataTable
            For Each tableName As String In tableNameList
                dt = ds.Tables(tableName)
                Dim sheetName As String = tableName
                'ヘッダー名称のリスト
                Dim headers As List(Of String) = New List(Of String)
                For Each col As System.Data.DataColumn In dt.Columns
                    headers.Add(col.ColumnName)
                Next
                xlsSheets = xlsBook.Worksheets

                '（注意）シートのインデックスは１から始まる
                xlsSheet = DirectCast(xlsSheets.Add, Excel.Worksheet)
                xlsSheet.Name = sheetName


                For i As Integer = 0 To headers.Count - 1
                    xlsRange = DirectCast(xlsSheet.Cells(1, i + 1), Excel.Range)
                    xlsRange.Value = headers.Item(i)
                Next

                ' セルに値を設定する。
                Dim sheetRowIndex As Integer = ROW_OFFSET

                For Each row As DataRow In dt.Rows
                    Dim sheetColumnIndex As Integer = COLUMN_OFFSET

                    For Each column As DataColumn In dt.Columns

                        If Not row.IsNull(column) Then
                            xlsRange = DirectCast(xlsSheet.Cells(sheetRowIndex, sheetColumnIndex), Excel.Range)

                            If column.DataType.Name = "Integer" Or _
                               column.DataType.Name = "Int32" Or _
                               column.DataType.Name = "Decimal" Or _
                                column.DataType.Name = "Long" Or _
                                column.DataType.Name = "Double" Or _
                                column.DataType.Name = "Short" Then
                                'セルの書式を数値型に設定
                                xlsRange.NumberFormatLocal = "G/標準"
                            ElseIf column.DataType.Name = "DateTime" Then
                                xlsRange.NumberFormatLocal = "yyyy/m/d h:mm"
                            Else
                                'セルの書式を文字列型に設定
                                xlsRange.NumberFormatLocal = "@"
                            End If

                            xlsRange.Value = row(column)
                            ReleaseComObject(DirectCast(xlsRange, Object))
                            sheetColumnIndex += 1
                        End If
                    Next
                    sheetRowIndex += 1
                Next

            Next
            xlsSheet = DirectCast(xlsSheets("Sheet1"), Excel.Worksheet)
            xlsSheet.Delete()

            ' 保存
            xlsBook.Save()



            Return True
        Catch ex As Exception
            Return False
        Finally
            'エクセル関係のオブジェクトは必ず解放すること
            ReleaseComObject(DirectCast(xlsRange, Object))
            ReleaseComObject(DirectCast(xlsSheet, Object))
            ReleaseComObject(DirectCast(xlsSheets, Object))
            xlsBook.Close(False)
            ReleaseComObject(DirectCast(xlsBook, Object))
            xlsApplication.Quit()
            ReleaseComObject(DirectCast(xlsApplication, Object))
        End Try

        Return False
    End Function



    '''<summary>
    '''ハッシュテーブルをエクセルファイルに出力します
    ''' Hashtable：key=SheetName,value=DataTable
    '''</summary>
    Public Function CreateExcelFromHashTable(ByVal ht As Dictionary(Of String, DataTable), _
                                                  ByVal filePath As String) As Boolean
        If IO.File.Exists(filePath) Then
            IO.File.Delete(filePath)
        End If


        Dim xlsApplication As New Excel.Application
        Dim xlsBook As Excel.Workbook = Nothing
        Dim xlsSheets As Excel.Sheets = Nothing
        Dim xlsSheet As Excel.Worksheet = Nothing
        Dim xlsRange As Excel.Range = Nothing

        xlsApplication.DisplayAlerts = False   '保存時の確認ダイアログを表示しない
        xlsBook = xlsApplication.Workbooks.Add()
        Dim tableNameList As New List(Of String)

        Try

            For Each sheetName As String In ht.Keys
                Dim dt As DataTable = DirectCast(ht(sheetName), DataTable)
                'ヘッダー名称のリスト
                Dim headers As List(Of String) = New List(Of String)
                For Each col As System.Data.DataColumn In dt.Columns
                    headers.Add(col.ColumnName)
                Next
                xlsSheets = xlsBook.Worksheets

                '（注意）シートのインデックスは１から始まる
                xlsSheet = DirectCast(xlsSheets.Add, Excel.Worksheet)
                xlsSheet.Name = sheetName


                For i As Integer = 0 To headers.Count - 1
                    xlsRange = DirectCast(xlsSheet.Cells(1, i + 1), Excel.Range)
                    xlsRange.Value = headers.Item(i)
                Next

                ' セルに値を設定する。
                Dim sheetRowIndex As Integer = ROW_OFFSET

                For Each row As DataRow In dt.Rows
                    Dim sheetColumnIndex As Integer = COLUMN_OFFSET

                    For Each column As DataColumn In dt.Columns

                        If Not row.IsNull(column) Then
                            xlsRange = DirectCast(xlsSheet.Cells(sheetRowIndex, sheetColumnIndex), Excel.Range)

                            If column.DataType.Name = "Integer" Or _
                               column.DataType.Name = "Int32" Or _
                               column.DataType.Name = "Decimal" Or _
                                column.DataType.Name = "Long" Or _
                                column.DataType.Name = "Double" Or _
                                column.DataType.Name = "Short" Then
                                'セルの書式を数値型に設定
                                xlsRange.NumberFormatLocal = "G/標準"
                            ElseIf column.DataType.Name = "DateTime" Then
                                xlsRange.NumberFormatLocal = "yyyy/m/d h:mm"
                            Else
                                'セルの書式を文字列型に設定
                                xlsRange.NumberFormatLocal = "@"
                            End If

                            xlsRange.Value = row(column)
                            ReleaseComObject(DirectCast(xlsRange, Object))
                            sheetColumnIndex += 1
                        End If
                    Next
                    sheetRowIndex += 1
                Next

            Next
            xlsSheet = DirectCast(xlsSheets("Sheet1"), Excel.Worksheet)
            xlsSheet.Delete()

            ' 保存
            xlsBook.Save()



            Return True
        Catch ex As Exception
            Return False
        Finally
            'エクセル関係のオブジェクトは必ず解放すること
            ReleaseComObject(DirectCast(xlsRange, Object))
            ReleaseComObject(DirectCast(xlsSheet, Object))
            ReleaseComObject(DirectCast(xlsSheets, Object))
            xlsBook.Close(False)
            ReleaseComObject(DirectCast(xlsBook, Object))
            xlsApplication.Quit()
            ReleaseComObject(DirectCast(xlsApplication, Object))
        End Try



        'IO.File.Move(tempFile, filePath)


        Return False
    End Function

    '''<summary>
    '''データセットをエクセルファイルに出力します
    '''</summary>
    Public Function CreateExcelFromDataSet(ByVal ds As DataSet, _
                                                  ByVal filePath As String, _
                                                  ByVal sheetName As String) As Boolean



        Dim xlsApplication As New Excel.Application
        Dim xlsBooks As Excel.Workbooks = Nothing
        Dim xlsBook As Excel.Workbook = Nothing
        Dim xlsSheets As Excel.Sheets = Nothing
        Dim xlsSheet As Excel.Worksheet = Nothing
        Dim xlsRange As Excel.Range = Nothing

        xlsApplication.DisplayAlerts = False   '保存時の確認ダイアログを表示しない

        Try

            For Each dt As DataTable In ds.Tables

                'ヘッダー名称のリスト
                Dim headers As List(Of String) = New List(Of String)
                For Each col As System.Data.DataColumn In dt.Columns
                    headers.Add(col.ColumnName)
                Next
                xlsBooks = xlsApplication.Workbooks
                xlsBook = xlsBooks.Add
                xlsSheets = xlsBook.Worksheets

                '（注意）シートのインデックスは１から始まる
                If xlsSheets.Count = 1 Then
                    xlsSheet = DirectCast(xlsSheets.Item(1), Excel.Worksheet)
                    xlsSheet.Name = sheetName
                Else
                    xlsSheet = DirectCast(xlsSheets.Add, Excel.Worksheet)
                    xlsSheet.Name = sheetName
                End If

                For i As Integer = 0 To headers.Count - 1
                    xlsRange = DirectCast(xlsSheet.Cells(1, i + 1), Excel.Range)
                    xlsRange.Value = headers.Item(i)
                Next

                ' セルに値を設定する。
                Dim sheetRowIndex As Integer = ROW_OFFSET

                For Each row As DataRow In dt.Rows
                    Dim sheetColumnIndex As Integer = COLUMN_OFFSET

                    For Each column As DataColumn In dt.Columns

                        If Not row.IsNull(column) Then
                            xlsRange = DirectCast(xlsSheet.Cells(sheetRowIndex, sheetColumnIndex), Excel.Range)

                            If column.DataType.Name = "Integer" Or _
                               column.DataType.Name = "Int32" Or _
                               column.DataType.Name = "Decimal" Or _
                                column.DataType.Name = "Long" Or _
                                column.DataType.Name = "Double" Or _
                                column.DataType.Name = "Short" Then
                                'セルの書式を数値型に設定
                                xlsRange.NumberFormatLocal = "G/標準"
                            ElseIf column.DataType.Name = "DateTime" Then
                                xlsRange.NumberFormatLocal = "yyyy/m/d h:mm"
                            Else
                                'セルの書式を文字列型に設定
                                xlsRange.NumberFormatLocal = "@"
                            End If

                            xlsRange.Value = row(column)
                            ReleaseComObject(DirectCast(xlsRange, Object))
                            sheetColumnIndex += 1
                        End If
                    Next
                    sheetRowIndex += 1
                Next

            Next

            ' 保存
            xlsBook.SaveAs(filePath)

            Return True

        Catch ex As Exception
            Return False
        Finally
            'エクセル関係のオブジェクトは必ず解放すること
            ReleaseComObject(DirectCast(xlsRange, Object))
            ReleaseComObject(DirectCast(xlsSheet, Object))
            ReleaseComObject(DirectCast(xlsSheets, Object))
            xlsBook.Close(False)
            ReleaseComObject(DirectCast(xlsBook, Object))
            ReleaseComObject(DirectCast(xlsBooks, Object))
            xlsApplication.Quit()
            ReleaseComObject(DirectCast(xlsApplication, Object))
        End Try

        Return False
    End Function

    '''<summary>
    '''データテーブルをエクセルファイルに出力します
    '''</summary>
    Public Function CreateExcelFromDataTable(ByVal dt As DataTable, _
                                                  ByVal filePath As String, _
                                                  ByVal sheetName As String) As Boolean
        'ヘッダー名称のリスト
        Dim headers As List(Of String) = New List(Of String)

        For Each col As System.Data.DataColumn In dt.Columns
            headers.Add(col.ColumnName)
        Next

        Dim xlsApplication As New Excel.Application
        Dim xlsBooks As Excel.Workbooks = Nothing
        Dim xlsBook As Excel.Workbook = Nothing
        Dim xlsSheets As Excel.Sheets = Nothing
        Dim xlsSheet As Excel.Worksheet = Nothing
        Dim xlsRange As Excel.Range = Nothing

        Try
            xlsApplication.DisplayAlerts = False   '保存時の確認ダイアログを表示しない

            xlsBooks = xlsApplication.Workbooks
            xlsBook = xlsBooks.Add
            xlsSheets = xlsBook.Worksheets

            '（注意）シートのインデックスは１から始まる
            xlsSheet = DirectCast(xlsSheets.Item(1), Excel.Worksheet)
            xlsSheet.Name = sheetName


            For i As Integer = 0 To headers.Count - 1
                xlsRange = DirectCast(xlsSheet.Cells(1, i + 1), Excel.Range)
                xlsRange.Value = headers.Item(i)
            Next

            ' セルに値を設定する。
            Dim sheetRowIndex As Integer = ROW_OFFSET

            For Each row As DataRow In dt.Rows
                Dim sheetColumnIndex As Integer = COLUMN_OFFSET

                For Each column As DataColumn In dt.Columns

                    If Not row.IsNull(column) Then
                        xlsRange = DirectCast(xlsSheet.Cells(sheetRowIndex, sheetColumnIndex), Excel.Range)

                        If column.DataType.Name = "Integer" Or _
                           column.DataType.Name = "Int32" Or _
                           column.DataType.Name = "Decimal" Or _
                            column.DataType.Name = "Long" Or _
                            column.DataType.Name = "Double" Or _
                            column.DataType.Name = "Short" Then
                            'セルの書式を数値型に設定
                            xlsRange.NumberFormatLocal = "G/標準"
                        ElseIf column.DataType.Name = "DateTime" Then
                            xlsRange.NumberFormatLocal = "yyyy/m/d h:mm"
                        Else
                            'セルの書式を文字列型に設定
                            xlsRange.NumberFormatLocal = "@"
                        End If

                        xlsRange.Value = row(column)
                        ReleaseComObject(DirectCast(xlsRange, Object))
                        sheetColumnIndex += 1
                    End If
                Next
                sheetRowIndex += 1
            Next

            ' 保存
            xlsBook.SaveAs(filePath)

            Return True

        Catch ex As Exception
            Return False
        Finally
            'エクセル関係のオブジェクトは必ず解放すること
            ReleaseComObject(DirectCast(xlsRange, Object))
            ReleaseComObject(DirectCast(xlsSheet, Object))
            ReleaseComObject(DirectCast(xlsSheets, Object))
            xlsBook.Close(False)
            ReleaseComObject(DirectCast(xlsBook, Object))
            ReleaseComObject(DirectCast(xlsBooks, Object))
            xlsApplication.Quit()
            ReleaseComObject(DirectCast(xlsApplication, Object))
        End Try

        Return False
    End Function

    ''' <summary>
    ''' COMオブジェクトを開放します。
    ''' </summary>
    Private Sub ReleaseComObject(ByRef target As Object)
        Try
            If Not target Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(target)
            End If
        Finally
            target = Nothing
        End Try
    End Sub
End Module
