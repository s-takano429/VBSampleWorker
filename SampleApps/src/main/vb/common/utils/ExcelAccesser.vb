Imports System.Data.OleDb
Imports System.Text.RegularExpressions

Public Class ExcelAccesser
    ''' <summary>
    ''' DBに対して実行するSQLステートメントを表します。
    ''' </summary>
    Private _oCommand As OleDbCommand
    Private _oConn As OleDbConnection

    Private _isConnected As Boolean = False


    Private Shared _accesser As ExcelAccesser = New ExcelAccesser()

    Private _adapterDict As New Dictionary(Of String, OleDbDataAdapter)

    ' コンストラクタです。(外部からのアクセス不可)
    Private Sub New()
        'Do nothing
    End Sub

    ''' <summary>
    ''' インスタンスを取得します。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getInstance() As ExcelAccesser
        Return _accesser
    End Function

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
    ''' DBへ接続します。
    ''' </summary>
    ''' <param name="path">DBファイルパス</param>
    ''' <param name="pass">DBパスワード</param>
    ''' <param name="exclusiveMode">排他モード</param>
    ''' <remarks></remarks>
    Public Sub dbConn(ByRef path As String, Optional ByVal pass As String = Nothing, Optional ByVal exclusiveMode As Boolean = False)
        Dim oConn As New OleDbConnection
        If Not _isConnected Then

            Try
                With New OleDbConnectionStringBuilder()
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                    '.Provider = "Microsoft.JET.OLEDB.4.0"
                    .DataSource = path
                    .Item("Extended Properties") = "Excel 12.0;HDR=YES"

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
                _isConnected = True
            Catch ex As Exception
                Throw New DbAccessException(ex.ToString)
            End Try
        End If

    End Sub

    ''' <summary>
    ''' DBを切断します。
    ''' </summary>
    Public Sub dbClose()
        Dim oConn As New OleDbConnection()
        If _isConnected Then

            Try
                oConn = OCommand.Connection
                If Not oConn Is Nothing Then
                    oConn.Close()
                End If
                _isConnected = False
            Catch ex As Exception
                Throw New DbAccessException(ex.ToString)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' DBからデータを取得します。（DB接続済み）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得データセット</returns>
    Public Function fillDataSet(ByRef sql As String) As DataSet
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
    ''' DBからデータを取得する（結果セットの最初の行の最初の列。結果セットが空の場合は、null 参照。）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得DataTable</returns>
    Public Function getCellData(ByVal sql As String) As Object
        Try
            OCommand.CommandText = sql
            Return OCommand.ExecuteScalar
        Catch ex As Exception
            Throw New DbAccessException(ex.ToString)
        End Try
    End Function

    ''' <summary>
    ''' DBからデータを取得する（DataTable）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得DataTable</returns>
    Public Function getDataTable(ByVal sql As String) As DataTable
        Dim oDataAdapter As New OleDbDataAdapter
        Dim oDataTable As New DataTable

        OCommand.CommandText = sql
        oDataAdapter.SelectCommand = OCommand
        oDataAdapter.Fill(oDataTable)

        Return oDataTable
    End Function

    ''' <summary>
    ''' DBにクエリを実行する
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    Public Sub execSQL(ByVal sql As String)
        Dim oDataAdapter As New OleDbDataAdapter
        Dim oDataTable As New DataTable

        OCommand.CommandText = sql
        OCommand.ExecuteNonQuery()

    End Sub

    ''' <summary>
    ''' DBからデータを取得する（DataTable）
    ''' </summary>
    ''' <param name="sql">SQLステートメント</param>
    ''' <returns>取得DataTable</returns>
    Public Function getUpdatableDataTable(ByVal sql As String, ByVal tableName As String) As DataTable
        Dim oDataTable As New DataTable

        Dim adapter As OleDbDataAdapter


        If _adapterDict.ContainsKey(tableName) Then
            adapter = _adapterDict.Item(tableName)
        Else
            adapter = New OleDbDataAdapter
            _adapterDict.Add(tableName, adapter)
        End If

        OCommand.CommandText = sql
        adapter.SelectCommand = OCommand
        adapter.Fill(oDataTable)


        Return oDataTable
    End Function




    Public Sub updateDataTable(ByVal dt As DataTable, ByVal tableName As String)
        _adapterDict.Item(tableName).Update(dt)
    End Sub





End Class
