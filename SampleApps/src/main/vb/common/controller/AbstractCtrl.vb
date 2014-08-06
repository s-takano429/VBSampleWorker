Public MustInherit Class AbstractCtrl

    '複数バックグランド処理機能時の注意文
    Private Const IS_BUSY As String = "既に別の機能が起動中です。" & vbCrLf & "終了するまで待つか、処理をキャンセルしてください。"

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

    'バックグランド処理オブジェクト
    Friend WithEvents backGroundWorker As New System.ComponentModel.BackgroundWorker

    'ワーカーによる処理か
    Private isWorkerProcess As Boolean = False

    ''' <summary>
    ''' 初期値の設定を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overridable Sub initialize()
        'Do nothing.
    End Sub

    ''' <summary>
    ''' 入力値の検証を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public MustOverride Sub validate()

    ''' <summary>
    ''' 処理を行うオブジェクトのオープン処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public MustOverride Sub open()

    ''' <summary>
    ''' 処理を行うオブジェクトの実行処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public MustOverride Sub execute()

    ''' <summary>
    ''' 処理を行うオブジェクトのクローズ処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public MustOverride Sub close()

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
    ''' 一連の処理を一括で行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングは行わず一括で上位メソッドに例外を飛ばす。</remarks>
    Public Overridable Sub run()
        Dim ex As Exception
        Try
            initialize()
            validate()
            open()
            execute()
        Catch executionExp As Exception
            Debug.WriteLine(executionExp.Message)
            Debug.WriteLine(executionExp.StackTrace)
            If isPrintingStackTrace Then
                MessageBox.Show(executionExp.Message & vbCrLf & vbCrLf & executionExp.StackTrace)
            Else
                MessageBox.Show(executionExp.Message)
            End If
            ex = executionExp
        Finally
            Try
                close()
            Catch closeExp As Exception
                Debug.WriteLine(closeExp.Message)
                Debug.WriteLine(closeExp.StackTrace)
                If isPrintingStackTrace Then
                    MessageBox.Show(closeExp.Message & vbCrLf & vbCrLf & closeExp.StackTrace)
                Else
                    MessageBox.Show(closeExp.Message)
                End If
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' 一連の処理をバックグランド処理で一括で行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングは行わず一括で上位メソッドに例外を飛ばす。</remarks>
    Public Sub startWorker()
        initialize()
        validate()

        If backGroundWorker.IsBusy Then
            MessageBox.Show(IS_BUSY, "処理中")
            Exit Sub
        End If
        isWorkerProcess = True
        backGroundWorker.WorkerSupportsCancellation = True
        backGroundWorker.WorkerReportsProgress = True
        backGroundWorker.RunWorkerAsync()
    End Sub

    ''' <summary>
    ''' バックグランド処理をキャンセルする。
    ''' </summary>
    ''' <remarks>例外のハンドリングは行わず一括で上位メソッドに例外を飛ばす。</remarks>
    Public Sub stopWorker()
        backGroundWorker.CancelAsync()
    End Sub

    ''' <summary>
    ''' 通知用オブジェクトの設定。
    ''' </summary>
    ''' <param name="mdl"></param>
    ''' <remarks></remarks>
    Protected Sub setProgressToModel(ByVal mdl As AbstractMdl)
        Dim progressObj As New ProgressFormControl
        If isWorkerProcess Then
            progressObj.backGroundWorker = backGroundWorker
        Else
            progressObj = ctrlProgressObject
        End If
        mdl.ctrlProgressObject = progressObj
    End Sub

    ''' <summary>
    ''' 処理のキャンセルが行われているかを確認する。
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub checkCancellation()
        If backGroundWorker.CancellationPending Then
            Throw New Exception("実行がキャンセルされました。")
        End If
    End Sub


    ''' <summary>
    ''' プログレステキストを変更する。
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <remarks></remarks>
    Protected Sub changeProgressText(ByVal txt As String)
        If isWorkerProcess Then
            backGroundWorker.ReportProgress(0, txt)
        Else
            ctrlProgressObject.changeProgressText(txt)
        End If
    End Sub

    ''' <summary>
    ''' プログレスバーの値を変更する。
    ''' </summary>
    ''' <param name="progressValue"></param>
    ''' <remarks></remarks>
    Protected Sub changeProgressValue(ByVal progressValue As Integer)
        If isWorkerProcess Then
            backGroundWorker.ReportProgress(progressValue, "")
        Else
            ctrlProgressObject.changeProgressValue(progressValue)
        End If
    End Sub

    ''' <summary>
    ''' プログレステキストとバーの値を変更する。
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <param name="progressValue"></param>
    ''' <remarks></remarks>
    Protected Sub changeProgress(ByVal txt As String, ByVal progressValue As Integer)
        If isWorkerProcess Then
            backGroundWorker.ReportProgress(progressValue, txt)
        Else
            ctrlProgressObject.changeProgress(txt, progressValue)
        End If
    End Sub




    ''' <summary>
    ''' バックグランド処理。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub backGroundWorker_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backGroundWorker.DoWork
        Dim ex As Exception = Nothing
        Try
            open()
            execute()
        Catch executionExp As Exception
            Debug.WriteLine(executionExp.Message)
            Debug.WriteLine(executionExp.StackTrace)
            If isPrintingStackTrace Then
                MessageBox.Show(executionExp.Message & vbCrLf & vbCrLf & executionExp.StackTrace)
            Else
                MessageBox.Show(executionExp.Message)
            End If
            ex = executionExp
        Finally
            Try
                close()
            Catch closeExp As Exception
                Debug.WriteLine(closeExp.Message)
                Debug.WriteLine(closeExp.StackTrace)
                If isPrintingStackTrace Then
                    MessageBox.Show(closeExp.Message & vbCrLf & vbCrLf & closeExp.StackTrace)
                Else
                    MessageBox.Show(closeExp.Message)
                End If
                ex = closeExp
            End Try
        End Try
        If Not IsNothing(ex) Then
            e.Cancel = True
        End If

    End Sub



    ''' <summary>
    ''' バックグランド処理終了時処理。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub backGroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backGroundWorker.RunWorkerCompleted
        If e.Cancelled Then
            cancelled()
        Else
            completed()
        End If
    End Sub

    ''' <summary>
    ''' 進行状態の変更あった場合。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub backGroundWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles backGroundWorker.ProgressChanged
        If e.ProgressPercentage = 0 Then
            ctrlProgressObject.changeProgressText(e.UserState.ToString)
        ElseIf e.UserState.ToString.Equals("") Then
            ctrlProgressObject.changeProgressValue(e.ProgressPercentage)
        Else
            ctrlProgressObject.changeProgress(e.UserState.ToString, e.ProgressPercentage)
        End If
    End Sub






End Class
