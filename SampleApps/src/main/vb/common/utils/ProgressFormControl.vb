Public Class ProgressFormControl
    'プログレスバーラベル
    Property progressLabel As ToolStripStatusLabel = New ToolStripStatusLabel

    'プログレスバー
    Property progressBar As ToolStripProgressBar = New ToolStripProgressBar

    'バックグランド処理オブジェクト
    Property backGroundWorker As System.ComponentModel.BackgroundWorker

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="label">プログレスバーラベル</param>
    ''' <param name="bar">プログレスバー</param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByRef label As ToolStripStatusLabel = Nothing, Optional ByRef bar As ToolStripProgressBar = Nothing)
        If IsNothing(label) = False Then
            progressLabel = label
        End If
        If IsNothing(bar) = False Then
            progressBar = bar
        End If
    End Sub

    ''' <summary>
    ''' プログレステキストを変更する。
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <remarks></remarks>
    Public Sub changeProgressText(ByVal txt As String)
        changeProgress(txt, progressBar.Value)
    End Sub

    ''' <summary>
    ''' プログレスバーの値を変更する。
    ''' </summary>
    ''' <param name="progressValue"></param>
    ''' <remarks></remarks>
    Public Sub changeProgressValue(ByVal progressValue As Integer)
        changeProgress(progressLabel.Text, progressValue)
    End Sub

    ''' <summary>
    ''' プログレステキストとバーの値を変更する。
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <param name="progressValue"></param>
    ''' <remarks></remarks>
    Public Sub changeProgress(ByVal txt As String, ByVal progressValue As Integer)
        If IsNothing(backGroundWorker) Then
            progressBar.Value = progressValue
            progressLabel.Text = txt
            Application.DoEvents()
        Else
            Try
                backGroundWorker.ReportProgress(progressValue, txt)
            Catch invalidOperationEx As InvalidOperationException
                Console.WriteLine("親スレッドが終了しているため子スレッドからコンポーネントの変更ができません。")
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub
End Class
