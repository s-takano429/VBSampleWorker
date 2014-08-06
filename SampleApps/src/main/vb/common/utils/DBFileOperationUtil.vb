Module DBFileOperationUtil

    Private Const ERRMSG_COPY_TEMPLATE_DB As String = "MDBテンプレートファイルが見つかりません。"

    ''' <summary>
    ''' テンプレートDBコピー
    ''' </summary>
    ''' <param name="SourceDBFilePath">テンプレートDBパス</param>
    ''' <param name="DestinationDBFilePath">出力先DBパス</param>
    ''' <remarks></remarks>
    Public Sub CopyDB(ByVal SourceDBFilePath As String, ByVal DestinationDBFilePath As String)

        'テンプレート存在確認
        If IO.File.Exists(SourceDBFilePath) = False Then
            Throw New Exception(ERRMSG_COPY_TEMPLATE_DB)
        End If

        '出力先に同名のファイルが存在する場合、消す
        If IO.File.Exists(DestinationDBFilePath) Then
            IO.File.Delete(DestinationDBFilePath)
        End If

        'テンプレートファイルをコピー
        IO.File.Copy(SourceDBFilePath, DestinationDBFilePath, True)

        '読み取り専用属性削除
        System.IO.File.SetAttributes(DestinationDBFilePath, IO.FileAttributes.Normal)

    End Sub


End Module
