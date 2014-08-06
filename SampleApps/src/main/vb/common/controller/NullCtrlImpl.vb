''' <summary>
''' 何も行わないController実装クラス.
''' </summary>
''' <remarks>
''' 実装スケルトンコード。
''' 初期化時などに使用し、Nullで処理が停止しないようにする。
''' </remarks>
Public Class NullCtrlImpl : Inherits AbstractCtrl




    ''' <summary>
    ''' 初期値の設定を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overrides Sub initialize()
        'Do nothing.
    End Sub

    ''' <summary>
    ''' 入力値の検証を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overrides Sub validate()
        'Do nothing.
    End Sub

    ''' <summary>
    ''' 処理を行うオブジェクトのオープン処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overrides Sub open()
        'Do nothing.
    End Sub

    ''' <summary>
    ''' 処理を行うオブジェクトの実行処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overrides Sub execute()
        'Do nothing.
    End Sub

    ''' <summary>
    ''' 処理を行うオブジェクトのクローズ処理を行う。
    ''' </summary>
    ''' <remarks>例外のハンドリングを行う。</remarks>
    Public Overrides Sub close()
        'Do nothing.
    End Sub



End Class