''' <summary>
''' Execute処理時例外のラッピングクラス
''' </summary>
''' <remarks></remarks>
Public Class ExecuteException : Inherits System.Exception
    Private Const ERROR_HEADER As String = "[ExecuteException]"

    Public Sub New(ByVal errorLog As String)
        MyBase.New(ERROR_HEADER & errorLog)
    End Sub
End Class
