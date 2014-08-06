Imports System.Data.OleDb
Imports System.Text.RegularExpressions

''' <summary>
''' DBへのアクセスを行ないます。
''' </summary>
Public Class DbAccessUtils

    ''' <summary>
    ''' DBに対して実行するSQLステートメントを表します。
    ''' </summary>
    Private _oCommand As OleDbCommand
    Private _oConn As OleDbConnection

    ''' <summary>
    ''' DBに対して実行するSQLステートメントを取得または設定します。
    ''' </summary>
    ''' <value>DBに対して実行するSQLステートメント</value>
    ''' <returns>DBに対して実行するSQLステートメント</returns>
    Private Property OCommand() As OleDbCommand
        Get
            Return _oCommand
        End Get
        Set(ByVal value As OleDbCommand)
            _oCommand = value
        End Set
    End Property

    ''' <summary>
    ''' 接続中のOleDbConnectionを公開します。
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DbConnection() As OleDbConnection
        Get
            Return _oConn
        End Get
    End Property

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    Public Sub New()
        _oCommand = New OleDbCommand
    End Sub

    ''' <summary>
    ''' DBへ接続します。
    ''' </summary>
    ''' <param name="path">DBファイルパス</param>
    ''' <param name="pass">DBパスワード</param>
    ''' <param name="exclusiveMode">排他モード</param>
    ''' <remarks></remarks>
    Public Sub DbConn(ByRef path As String, Optional ByVal pass As String = Nothing, Optional ByVal exclusiveMode As Boolean = False)
        Dim oConn As New OleDbConnection

        Try
            With New OleDbConnectionStringBuilder()
                .Provider = "Microsoft.JET.OLEDB.4.0"
                .DataSource = path

                If pass IsNot Nothing Then
                    .Item("Jet OLEDB:Database Password") = pass
                End If

                If exclusiveMode Then
                    .Item("Mode") = "Share Exclusive" '排他モードで開く
                End If

                oConn.ConnectionString = .ConnectionString
            End With
            oConn.Open()
            OCommand = oConn.CreateCommand
            OCommand.Connection = oConn
            _oConn = oConn
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' MDBに含まれるテーブル名一覧を取得(それぞれのテーブル名には、カラム名一覧が付随する)
    ''' </summary>
    ''' <returns>テーブル名リスト</returns>
    ''' <remarks></remarks>
    Public Function GetTableNameList() As Dictionary(Of String, Dictionary(Of String, String))
        Dim listTableName As New Dictionary(Of String, Dictionary(Of String, String))

        Try
            Dim table As DataTable = OCommand.Connection.GetSchema("Tables")
            If table Is Nothing Then
                Return listTableName
            End If

            For Each dr As DataRow In table.Rows
                For Each col As DataColumn In table.Columns
                    Console.WriteLine("{0} = {1}", col.ColumnName, dr(col))
                Next
                'テーブル以外は弾く
                If CStr(dr("TABLE_TYPE")).Equals("TABLE") = False Then Continue For

                'テーブル名取得
                Dim tableName As String = CStr(dr("TABLE_NAME"))

                '列情報取得   
                Dim dtColumns As DataTable = OCommand.Connection.GetOleDbSchemaTable( _
                 OleDbSchemaGuid.Columns, _
                 New Object() {Nothing, Nothing, tableName, Nothing})

                Dim listColumn As New Dictionary(Of String, String)
                For Each dr2 As DataRow In dtColumns.Rows
                    Dim columnName As String = CStr(dr2("COLUMN_NAME"))
                    listColumn.Add(columnName, columnName)
                Next

                'テーブル名リストに追加
                listTableName.Add(tableName, listColumn)
            Next
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
        End Try

        Return listTableName
    End Function

    ''' <summary>
    ''' DBへ接続します(XLS)。
    ''' </summary>
    ''' <param name="path">接続先DBパス</param>
    ''' <remarks></remarks>
    Public Sub DbConnXls(ByRef path As String)
        Dim oConn As New OleDbConnection

        Try
            With New OleDbConnectionStringBuilder()
                .Provider = "Microsoft.JET.OLEDB.4.0"
                .DataSource = path
                .Item("Extended Properties") = "Excel 8.0;HDR=YES;"
                oConn.ConnectionString = .ConnectionString
            End With
            oConn.Open()
            OCommand = oConn.CreateCommand
            OCommand.Connection = oConn
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' DBへ接続する（csv）
    ''' </summary>
    ''' <param name="path">接続先（ファイル名を含めない）</param>
    ''' <param name="hdr">ヘッダー有無 True=有 False=無</param>
    ''' <remarks></remarks>
    Public Sub OpenCsv(ByVal path As String, ByVal hdr As Boolean)
        Dim OConn As New OleDbConnection

        Try
            Dim hdrYesNo As String = "YES"
            If hdr = False Then
                hdrYesNo = "NO"
            End If

            With New OleDbConnectionStringBuilder()
                .Provider = "Microsoft.JET.OLEDB.4.0"
                .DataSource = path
                .Item("Extended Properties") = "Text;HDR=" & hdrYesNo & ";FMT=Delimited"
                OConn.ConnectionString = .ConnectionString
            End With

            'オープン
            OConn.Open()

            'OleDbCommandを保持する
            OCommand = OConn.CreateCommand
            OCommand.Connection = OConn
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' DBのパスワード変更
    ''' </summary>
    ''' <param name="path">変更するMDBファイル</param>
    ''' <param name="newPass">新しいパスワード</param>
    ''' <param name="oldPass">以前のパスワード</param>
    ''' <remarks></remarks>
    Public Sub DbChangePass(ByVal path As String, ByVal newPass As String, ByVal oldPass As String)
        Try
            DbConn(path, oldPass, True)

            If oldPass.Length = 0 Then
                oldPass = "NULL"
            End If
            ExecuteQuery("ALTER Database Password " + newPass + " " + oldPass)
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        Finally
            DbClose()
        End Try
    End Sub

    ''' <summary>
    ''' DBへ接続し、トランザクションを開始します。
    ''' </summary>
    ''' <param name="path">接続先DBパス</param>
    ''' <remarks></remarks>
    Public Sub DbConnTran(ByRef path As String)
        Dim oConn As OleDbConnection
        Dim oTran As OleDbTransaction

        Try
            DbConn(path)
            oConn = OCommand.Connection
            oTran = oConn.BeginTransaction
            OCommand.Transaction = oTran
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' トランザクションをコミットします。
    ''' </summary>
    Public Sub DbCommit()
        Dim oTran As OleDbTransaction

        Try
            oTran = OCommand.Transaction
            If Not oTran Is Nothing Then
                oTran.Commit()
            End If
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' トランザクションをロールバックします。
    ''' </summary>
    Public Sub DbRollback()
        Dim oTran As OleDbTransaction

        Try
            oTran = OCommand.Transaction
            If Not oTran Is Nothing Then
                oTran.Rollback()
            End If
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' DBを切断します。
    ''' </summary>
    Public Sub DbClose()
        Dim oConn As New OleDbConnection()

        Try
            oConn = OCommand.Connection
            If Not oConn Is Nothing Then
                oConn.Close()
            End If
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' DBからデータを取得します。（DB接続済み）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得データセット</returns>
    Public Function FillDataSet(ByRef sql As String) As DataSet
        Dim oDataAdapter As New OleDbDataAdapter
        Dim oDataSet As DataSet = New DataSet()

        Try
            OCommand.CommandText = sql
            oDataAdapter.SelectCommand = OCommand
            oDataAdapter.Fill(oDataSet)
            Return oDataSet
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' DBからデータを取得します。（DB接続）
    ''' </summary>
    ''' <param name="path">接続先DBパス</param>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得データセット</returns>
    Public Function FillDataSet(ByRef path As String, ByRef sql As String) As DataSet
        DbConn(path)

        Try
            Dim oDataSet As DataSet = FillDataSet(sql)
            DbClose()
            Return oDataSet
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' DBからデータを取得する（DataTable）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得DataTable</returns>
    Public Function GetDataTable(ByVal sql As String) As DataTable
        Dim oDataAdapter As New OleDbDataAdapter
        Dim oDataTable As New DataTable

        OCommand.CommandText = sql
        oDataAdapter.SelectCommand = OCommand
        oDataAdapter.Fill(oDataTable)

        Return oDataTable
    End Function

    ''' <summary>
    ''' DBからデータを取得する（結果セットの最初の行の最初の列。結果セットが空の場合は、null 参照。）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得DataTable</returns>
    Public Function GetCellData(ByVal sql As String) As Object
        Try
            OCommand.CommandText = sql
            Return OCommand.ExecuteScalar
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' クエリを実行します。（DB接続済み）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>クエリ実行結果</returns>
    Public Function ExecuteQuery(ByRef sql As String) As Integer
        Try
            OCommand.CommandText = sql
            Return OCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' クエリを実行します。（DB接続）
    ''' </summary>
    ''' <param name="path">接続先DBパス</param>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>クエリ実行結果</returns>
    Public Function ExecuteQuery(ByRef path As String, ByRef sql As String) As Integer
        DbConn(path)

        Try
            Dim ret As Integer = ExecuteQuery(sql)
            Return ret
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        Finally
            DbClose()
        End Try
    End Function

    ''' <summary>
    ''' Excelファイルからシート名を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSheetName() As String
        Dim oConn As OleDbConnection = OCommand.Connection

        Dim sheetName As String

        '左のシート名を取得
        sheetName = oConn.GetSchema("Tables").Rows(0).Item("TABLE_NAME").ToString

        '取得したシート名の末尾に'$'がついていない場合は、付与する
        If Microsoft.VisualBasic.Right(sheetName, 1) <> "$" Then
            sheetName += "$"
        End If

        Return sheetName
    End Function

    '    ''' <summary>
    '    ''' MDBファイルを最適化します。
    '    ''' </summary>
    '    ''' <param name="path">DBファイルパス</param>
    '    ''' <param name="pass">DBパスワード</param>
    '    Public Shared Sub CompactMdb(ByRef path As String, Optional ByVal pass As String = Nothing)
    '        Dim jroJet As JRO.JetEngine

    '        ' DB存在確認
    '        If Not IO.File.Exists(path) Then
    '            Throw New DbAccessException("データベースが見つかりません。")
    '        End If

    '        Try
    '            jroJet = New JRO.JetEngine

    '            ' 圧縮元
    '            Dim source As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '             "Data Source=" & path & ";"
    '            If pass IsNot Nothing Then
    '                source &= "Jet OLEDB:Database Password=" & pass
    '            End If

    '            ' 圧縮先
    '            Dim targetPath As String = path.Substring(0, path.LastIndexOf("."c)) & "_cmp.mdb"
    '            Dim target As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '             "Data Source=" & targetPath & ";" & _
    '             "Jet OLEDB:Engine Type=5;"
    '            If pass IsNot Nothing Then
    '                target &= "Jet OLEDB:Database Password=" & pass
    '            End If

    '            ' 圧縮
    '            jroJet.CompactDatabase(source, target)

    '            ' 圧縮元ファイル削除
    '            My.Computer.FileSystem.DeleteFile(path, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

    '            ' 圧縮先ファイルリネーム
    '            My.Computer.FileSystem.MoveFile(targetPath, path, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)

    '            'MessageBox.Show("圧縮しました。", "MDB圧縮", MessageBoxButtons.OK)

    '            jroJet = Nothing
    '        Catch ex As Exception
    '            jroJet = Nothing
    '            Throw New DbAccessException(ex.ToString)
    '        End Try
    '    End Sub
End Class

Public Class DbAccessException
    Inherits ApplicationException

    Public Sub New(ByVal errorMessage As String)
        MyBase.new(errorMessage)
    End Sub
End Class
