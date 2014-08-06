Public MustInherit Class AbstractMdl

    '初期化処理例外時の例外メッセージ
    Protected Const INITIALIZE_ERROR As String = "Initialize Error."

    'バリデーション処理例外時の例外メッセージ
    Protected Const VALIDATION_ERROR As String = "Validation Error."

    'オープン処理例外時の例外メッセージ
    Protected Const OPEN_ERROR As String = "Open Error."

    'クローズ処理例外時の例外メッセージ
    Protected Const CLOSE_ERROR As String = "Close Error."

    '実行処理例外時の例外メッセージ
    Protected Const EXECUTE_ERROR As String = "Execute Error."

    'stackTraceを表示するか
    Property isPrintingStackTrace As Boolean = False

    'ToolStripStatusをコントロールするクラス
    Property ctrlProgressObject As New ProgressFormControl

    '
    Private thread As System.Threading.Thread

    ''' <summary>
    ''' Thread内処理。
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub run()
        'Do nothing.
    End Sub



    ''' <summary>
    ''' Thread処理を開始する。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub startThread()
        thread = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf runThread))
        thread.IsBackground = True
        thread.Start()
    End Sub



    ''' <summary>
    ''' Thread処理を強制終了する。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub stopThread()
        thread.Abort()
    End Sub

    ''' <summary>
    ''' Thread処理の終了待機をする。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub joinThread()
        thread.Join()
    End Sub

    ''' <summary>
    ''' バックグランド処理終了時の処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overridable Sub completed()
        'Do nothing.
    End Sub

    ''' <summary>
    ''' バックグランド処理キャンセル時の処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overridable Sub cancelled()
        'Do nothing.
    End Sub
    ''' <summary>
    ''' プログレステキストを変更する。
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <remarks></remarks>
    Protected Sub changeProgressText(ByVal txt As String)
        ctrlProgressObject.changeProgressText(txt)
    End Sub

    ''' <summary>
    ''' プログレスバーの値を変更する。
    ''' </summary>
    ''' <param name="progressValue"></param>
    ''' <remarks></remarks>
    Protected Sub changeProgressValue(ByVal progressValue As Integer)
        ctrlProgressObject.changeProgressValue(progressValue)
    End Sub

    ''' <summary>
    ''' プログレステキストとバーの値を変更する。
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <param name="progressValue"></param>
    ''' <remarks></remarks>
    Protected Sub changeProgress(ByVal txt As String, ByVal progressValue As Integer)
        ctrlProgressObject.changeProgress(txt, progressValue)
    End Sub

    ''' <summary>
    ''' Thread実行メソッド。
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub runThread()
        Try
            run()
        Catch taEx As System.Threading.ThreadAbortException
            cancelled()
        Catch invalidOperationEx As InvalidOperationException
            Debug.WriteLine(invalidOperationEx.Message)
            Debug.WriteLine(invalidOperationEx.StackTrace)
            If isPrintingStackTrace Then
                Debug.WriteLine("親スレッドが終了しているため子スレッドからコンポーネントの変更ができません。" & vbCrLf & invalidOperationEx.StackTrace)
            Else
                Debug.WriteLine("親スレッドが終了しているため子スレッドからコンポーネントの変更ができません。")
            End If
        Catch ex As Exception
            Throw
        End Try
        completed()
    End Sub

End Class
