Public Class CtrlMainImpl
    : Inherits AbstractCtrl

    Property test As String

    Public Overrides Sub validate()
        'Do nothing.
    End Sub

    Public Overrides Sub open()
        'Do nothing.
    End Sub

    Public Overrides Sub execute()

        Dim s As String = "あいうえお546か453き886く462けこ"

        Console.WriteLine(StringUtils.reverse(StringUtils.right(s, 3)))

        Console.WriteLine(StringUtils.match(s, "[1-9][1-9][1-9]"))


        For Each ts As String In StringUtils.matches(s, "[1-9][1-9][1-9]")
            Console.WriteLine("test:" & ts)
        Next

        Console.WriteLine(StringUtils.removeEnd(s, 2))
        Console.WriteLine(StringUtils.removeStart(s, 2))

        For i As Integer = 0 To 100
            checkCancellation()
            Dim mdl1 As New MdlMain
            Dim mdl2 As New MdlMain
            Dim mdl3 As New MdlMain
            mdl1.text = " text 1"
            mdl2.text = " text 2"
            mdl3.text = " text 3"
            setProgressToModel(mdl1)
            setProgressToModel(mdl2)
            setProgressToModel(mdl3)
            mdl1.startThread()
            mdl2.startThread()
            mdl3.startThread()
            'Threading.Thread.Sleep(5000)
            'MessageBox.Show("sleep完了")
            mdl1.joinThread()
            mdl2.joinThread()
            mdl3.joinThread()
            'MessageBox.Show("完了")
        Next
    End Sub

    Public Overrides Sub close()
        'Do nothing.
    End Sub
End Class
