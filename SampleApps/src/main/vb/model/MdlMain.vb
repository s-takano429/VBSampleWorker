Public Class MdlMain
    : Inherits AbstractMdl
    Property text As String = ""



    Public Overrides Sub run()
        For i As Integer = 0 To 100
            changeProgress(text & " " & i & " カウント。", i)
            System.Threading.Thread.Sleep(10)
        Next
    End Sub


    Public Overrides Sub completed()
        MyBase.completed()
        MessageBox.Show("処理完了")
        changeProgressText("処理完了")
    End Sub
End Class
