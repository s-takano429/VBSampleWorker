''' <summary>
''' Initializ処理時例外のラッピングクラス
''' </summary>
''' <remarks></remarks>
Public Class InitializeException : Inherits System.Exception

    Private Const ERROR_HEADER As String = "[InitializeException]"

    Public Sub New(ByVal errorLog As String)
        MyBase.New(ERROR_HEADER & errorLog)
    End Sub

End Class
