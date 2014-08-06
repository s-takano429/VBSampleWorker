''' <summary>
''' Validate処理時例外のラッピングクラス
''' </summary>
''' <remarks></remarks>
Public Class ValidateException : Inherits System.Exception

    Private Const ERROR_HEADER As String = "[ValidateException]"

    Public Sub New(ByVal errorLog As String)
        MyBase.New(ERROR_HEADER & errorLog)
    End Sub

End Class
