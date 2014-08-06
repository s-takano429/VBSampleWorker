''' <summary>
''' Open処理時例外のラッピングクラス
''' </summary>
''' <remarks></remarks>
Public Class OpenException : Inherits System.Exception

    Private Const ERROR_HEADER As String = "[OpenException]"

    Public Sub New(ByVal errorLog As String)
        MyBase.New(ERROR_HEADER & errorLog)
    End Sub

End Class
