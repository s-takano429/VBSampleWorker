Module StringUtils
    ''' <summary>
    ''' 引数文字列が全て全角の場合にtrueを返します。
    ''' </summary>
    ''' <param name="str">チェック対象文字列</param>
    ''' <returns>引数文字列が全て全角の場合はtrue、そうでない場合はfalse</returns>
    Public Function isZenkakuStr(ByRef str As String) As Boolean
        If Not String.IsNullOrEmpty(str) Then
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
            Dim num As Integer = enc.GetByteCount(str)
            Return num = str.Length * 2
        End If
        Return False
    End Function

    ''' <summary>
    ''' 引数文字列が全て半角の場合にtrueを返します。
    ''' </summary>
    ''' <param name="str">チェック対象文字列</param>
    ''' <returns>引数文字列が全て半角の場合はtrue、そうでない場合はfalse</returns>
    Public Function isHankakuStr(ByRef str As String) As Boolean
        If Not String.IsNullOrEmpty(str) Then
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")
            Dim num As Integer = enc.GetByteCount(str)
            Return num = str.Length
        End If
        Return False
    End Function



    ''' <summary>
    ''' Shift-JISでテキストを読み込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <returns>ファイルを読み込んだリスト</returns>
    ''' <remarks></remarks>
    Public Function readTextShiftJIS(ByVal filePath As String) As List(Of String)
        Return readText(filePath, System.Text.Encoding.GetEncoding("shift_jis"))
    End Function

    ''' <summary>
    ''' UTF-8でテキストを読み込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <returns>ファイルを読み込んだリスト</returns>
    ''' <remarks></remarks>
    Public Function readTextUTF8(ByVal filePath As String) As List(Of String)
        Return readText(filePath, System.Text.Encoding.GetEncoding("UTF-8"))
    End Function

    ''' <summary>
    ''' UTF-8でテキストを読み込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <returns>ファイルを読み込んだリスト</returns>
    ''' <remarks></remarks>
    Public Function readText(ByVal filePath As String, ByVal encording As System.Text.Encoding) As List(Of String)
        Dim sr As New System.IO.StreamReader(filePath, encording)
        Dim src As New List(Of String)
        While sr.Peek() > -1
            src.Add(sr.ReadLine())
        End While
        Return src
    End Function


    ''' <summary>
    ''' Shift-JISでテキストを書き込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <remarks></remarks>
    Public Sub writeTextShiftJIS(ByVal filePath As String, ByVal writeString As String)
        writeText(filePath, System.Text.Encoding.GetEncoding("shift_jis"), writeString)
    End Sub

    ''' <summary>
    ''' UTF-8でテキストを書き込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <remarks></remarks>
    Public Sub writeTextUTF8(ByVal filePath As String, ByVal writeString As String)
        writeText(filePath, System.Text.Encoding.GetEncoding("UTF-8"), writeString)
    End Sub

    ''' <summary>
    ''' Shift-JISでテキストを書き込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <remarks></remarks>
    Public Sub writeTextShiftJIS(ByVal filePath As String, ByVal writeStringList As List(Of String))
        writeText(filePath, System.Text.Encoding.GetEncoding("shift_jis"), listToString(writeStringList))
    End Sub

    ''' <summary>
    ''' UTF-8でテキストを書き込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <remarks></remarks>
    Public Sub writeTextUTF8(ByVal filePath As String, ByVal writeStringList As List(Of String))
        writeText(filePath, System.Text.Encoding.GetEncoding("UTF-8"), listToString(writeStringList))
    End Sub

    ''' <summary>
    ''' リストから文字列に変換する。
    ''' </summary>
    ''' <param name="stringList">文字列リスト</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listToString(ByVal stringList As List(Of String)) As String
        Dim sBuilder As New System.Text.StringBuilder
        For Each line As String In stringList
            sBuilder.AppendLine(line)
        Next
        Return sBuilder.ToString
    End Function

    ''' <summary>
    ''' 文字列をファイル書き込む。
    ''' </summary>
    ''' <param name="filePath">ファイルパス</param>
    ''' <param name="encording">エンコード</param>
    ''' <param name="writeString">書き込み文字列</param>
    ''' <remarks>ファイルは上書きする</remarks>
    Public Sub writeText(ByVal filePath As String, ByVal encording As System.Text.Encoding, ByVal writeString As String)
        'ファイルを上書きし、Shift JISで書き込む 
        Dim sw As New System.IO.StreamWriter(filePath, False, encording)
        sw.Write(writeString)
        '閉じる
        sw.Close()
    End Sub

    ''' <summary>
    ''' 左から指定文字列抜出す。
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="len"></param>
    ''' <remarks></remarks>
    Public Function left(ByVal str As String, ByVal len As Integer) As String
        Return str.Substring(0, len)
    End Function

    ''' <summary>
    ''' 右から指定文字列抜出す。
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="len"></param>
    ''' <remarks></remarks>
    Public Function right(ByVal str As String, ByVal len As Integer) As String
        Return str.Substring(str.Length - len, len)
    End Function

    ''' <summary>
    ''' 検索対象文字列から正規表現パターンを抜出す。
    ''' </summary>
    ''' <param name="str">検索対象文字列</param>
    ''' <param name="regex">検索用正規表現</param>
    ''' <returns>検索一致文字列(複数一致があっても最初の1回の適合文字列のみ返す)</returns>
    ''' <remarks></remarks>
    Public Function match(ByVal str As String, ByVal regex As String) As String
        Dim ans As String = ""
        Dim pattern As New System.Text.RegularExpressions.Regex(regex, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim matchPattern As System.Text.RegularExpressions.Match = pattern.Match(str)

        If matchPattern.Success Then
            ans = matchPattern.Value
        End If
        Return ans
    End Function


    ''' <summary>
    ''' 検索対象文字列から正規表現パターンを抜出す。
    ''' </summary>
    ''' <param name="str">検索対象文字列</param>
    ''' <param name="regex">検索用正規表現</param>
    ''' <returns>検索一致文字列リスト</returns>
    ''' <remarks></remarks>
    Public Function matches(ByVal str As String, ByVal regex As String) As List(Of String)
        matches = New List(Of String)
        Dim pattern As New System.Text.RegularExpressions.Regex(regex, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim matchPattern As System.Text.RegularExpressions.MatchCollection = pattern.Matches(str)
        For Each match As System.Text.RegularExpressions.Match In matchPattern
            matches.Add(match.Value)
        Next

    End Function


    ''' <summary>
    ''' 反転文字列を返す。
    ''' </summary>
    ''' <param name="str">反転対象文字列</param>
    ''' <returns>反転文字</returns>
    ''' <remarks></remarks>
    Public Function reverse(ByVal str As String) As String
        Return StrReverse(str)
    End Function


    ''' <summary>
    ''' 末尾の指定文字数を削除した文字列を返す。
    ''' </summary>
    ''' <param name="str">対象文字列</param>
    ''' <param name="len">削除文字数</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function removeEnd(ByVal str As String, ByVal len As Integer) As String
        Return str.Remove(str.Length - len, len)
    End Function

    ''' <summary>
    ''' 初めの指定文字数を削除した文字列を返す。
    ''' </summary>
    ''' <param name="str">対象文字列</param>
    ''' <param name="len">削除文字数</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function removeStart(ByVal str As String, ByVal len As Integer) As String
        Return str.Remove(0, len)
    End Function

End Module
