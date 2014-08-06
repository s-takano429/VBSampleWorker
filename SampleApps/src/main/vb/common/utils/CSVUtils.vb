Module CSVUtils


    ''' <summary>
    ''' CSVファイルをDataTable（すべて文字列取込）に変換
    ''' </summary>
    ''' <param name="csvFilePath">CSVファイルパス</param>
    ''' <param name="HDR">ヘッダがあるか</param>
    ''' <returns>変換結果のDataTable</returns>
    ''' <remarks></remarks>
    Public Function CSVFileToDataTable(ByVal csvFilePath As String, Optional ByVal HDR As Boolean = True) As DataTable
        Dim sr As New System.IO.StreamReader(csvFilePath, System.Text.Encoding.GetEncoding("shift_jis"))
        Return CSVToDataTable(sr.ReadToEnd, HDR)
    End Function


    ''' <summary>
    ''' CSVテキストをDataTable（すべて文字列取込）に変換
    ''' </summary>
    ''' <param name="csvText">CSVの内容が入ったString</param>
    ''' <param name="HDR">ヘッダがあるか</param>
    ''' <returns>変換結果のDataTable</returns>
    ''' <remarks></remarks>
    Public Function CSVToDataTable(ByVal csvText As String, Optional ByVal HDR As Boolean = True) As DataTable
        '前後の改行を削除しておく
        csvText = csvText.Trim(New Char() {ControlChars.Cr, ControlChars.Lf})
        Dim isHDR As Boolean = True
        Dim csvRecords As New System.Collections.ArrayList
        Dim csvFields As New System.Collections.ArrayList

        Dim csvTextLength As Integer = csvText.Length
        Dim startPos As Integer = 0
        Dim endPos As Integer = 0
        Dim field As String = ""
        Dim dt As New DataTable

        While True
            '空白を飛ばす
            While startPos < csvTextLength _
                AndAlso (csvText.Chars(startPos) = " "c _
                OrElse csvText.Chars(startPos) = ControlChars.Tab)
                startPos += 1
            End While

            'データの最後の位置を取得
            If startPos < csvTextLength _
                AndAlso csvText.Chars(startPos) = ControlChars.Quote Then
                '"で囲まれているとき
                '最後の"を探す
                endPos = startPos
                While True
                    endPos = csvText.IndexOf(ControlChars.Quote, endPos + 1)
                    If endPos < 0 Then
                        Throw New ApplicationException("""が不正")
                    End If
                    '"が2つ続かない時は終了
                    If endPos + 1 = csvTextLength OrElse _
                        csvText.Chars((endPos + 1)) <> ControlChars.Quote Then
                        Exit While
                    End If
                    '"が2つ続く
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos + 1)
                '""を"にする
                field = field.Substring(1, field.Length - 2). _
                    Replace("""""", """")

                endPos += 1
                '空白を飛ばす
                While endPos < csvTextLength AndAlso _
                    csvText.Chars(endPos) <> ","c AndAlso _
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While
            Else
                '"で囲まれていない
                'カンマか改行の位置
                endPos = startPos
                While endPos < csvTextLength AndAlso _
                    csvText.Chars(endPos) <> ","c AndAlso _
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos)
                '後の空白を削除
                field = field.TrimEnd()
            End If


            'フィールドの追加
            csvFields.Add(field)

            '行の終了か調べる
            If endPos >= csvTextLength OrElse _
                csvText.Chars(endPos) = ControlChars.Lf Then
                '行の終了
                'レコードの追加
                csvFields.TrimToSize()
                csvRecords.Add(csvFields)

                If isHDR Then
                    For Each header As String In csvFields
                        dt.Columns.Add(header, GetType(System.String))
                    Next
                    isHDR = False
                Else
                    Dim row As DataRow = dt.NewRow
                    For index As Integer = 0 To csvFields.Count - 1
                        row(index) = csvFields(index)
                    Next
                    dt.Rows.Add(row)
                End If



                csvFields = New System.Collections.ArrayList(csvFields.Count)

                If endPos >= csvTextLength Then
                    '終了
                    Exit While
                End If
            End If

            '次のデータの開始位置
            startPos = endPos + 1
        End While

        csvRecords.TrimToSize()
        Return dt
    End Function


    ''' <summary>
    ''' DataTable型のデータをCSVファイルに出力する
    ''' </summary>
    ''' <param name="dt">出力するDataTable</param>
    ''' <param name="csvPath">出力ファイルパス</param>
    ''' <remarks></remarks>
    Public Sub exportDataTable(ByVal dt As DataTable, ByVal csvPath As String, Optional ByVal headerStr As String = Nothing)


        'CSVファイルに書き込むときに使うEncoding    
        Dim enc As System.Text.Encoding = _
            System.Text.Encoding.GetEncoding("Shift_JIS")


        '開く
        Dim sr As New System.IO.StreamWriter(csvPath, False, enc)

        Dim colCount As Integer = dt.Columns.Count
        Dim lastColIndex As Integer = colCount - 1

        'ヘッダを書き込む
        Dim i As Integer
        If headerStr Is Nothing Then
            For i = 0 To colCount - 1
                'ヘッダの取得
                Dim field As String = dt.Columns(i).Caption
                '"で囲む必要があるか調べる
                If field.IndexOf(ControlChars.Quote) > -1 OrElse _
                    field.IndexOf(","c) > -1 OrElse _
                    field.IndexOf(ControlChars.Cr) > -1 OrElse _
                    field.IndexOf(ControlChars.Lf) > -1 OrElse _
                    field.StartsWith(" ") OrElse _
                    field.StartsWith(ControlChars.Tab) OrElse _
                    field.EndsWith(" ") OrElse _
                    field.EndsWith(ControlChars.Tab) Then
                    If field.IndexOf(ControlChars.Quote) > -1 Then
                        '"を""とする
                        field = field.Replace("""", """""")
                    End If
                    field = """" + field + """"
                End If
                'フィールドを書き込む
                sr.Write(field)
                'カンマを書き込む
                If lastColIndex > i Then
                    sr.Write(","c)
                End If
            Next i
        Else
            sr.Write(headerStr)
        End If
        '改行する
        sr.Write(ControlChars.Cr + ControlChars.Lf)

        'レコードを書き込む
        Dim row As DataRow
        For Each row In dt.Rows
            For i = 0 To colCount - 1
                'フィールドの取得
                Dim field As String = row(i).ToString()
                '"で囲む必要があるか調べる
                If field.IndexOf(ControlChars.Quote) > -1 OrElse _
                    field.IndexOf(","c) > -1 OrElse _
                    field.IndexOf(ControlChars.Cr) > -1 OrElse _
                    field.IndexOf(ControlChars.Lf) > -1 OrElse _
                    field.StartsWith(" ") OrElse _
                    field.StartsWith(ControlChars.Tab) OrElse _
                    field.EndsWith(" ") OrElse _
                    field.EndsWith(ControlChars.Tab) Then
                    If field.IndexOf(ControlChars.Quote) > -1 Then
                        '"を""とする
                        field = field.Replace("""", """""")
                    End If
                    field = """" + field + """"
                End If
                'フィールドを書き込む
                sr.Write(field)
                'カンマを書き込む
                If lastColIndex > i Then
                    sr.Write(","c)
                End If
            Next i
            '改行する
            sr.Write(ControlChars.Cr + ControlChars.Lf)
        Next row

        '閉じる
        sr.Close()
    End Sub

End Module
