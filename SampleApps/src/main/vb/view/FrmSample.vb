

Public Class FrmSample

    Private ctrl As AbstractCtrl

    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = getToolsName()
        TSSBtnStop.Image = SystemIcons.Hand.ToBitmap
    End Sub

    Private Sub StopButton_ButtonClick(sender As Object, e As EventArgs) Handles TSSBtnStop.ButtonClick

        ctrl.stopWorker()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Test
        Dim ctrl As New CtrlMainImpl()
        ctrl.ctrlProgressObject = New ProgressFormControl(TSSLabel, TSProgressBar)
        ctrl.test = "tester"
        ctrl.startWorker()
        Me.ctrl = ctrl
    End Sub



End Class
