Public Class DateTimeUtils : Inherits Timer


    '初期のタイマー間隔
    Private Const DEFAULT_INTERVAL As Long = 10000

    '初期のタイマー間隔
    Private Const DEFAULT_DIV_TIME As Integer = 30

    '唯一のアクセス可能オブジェクト
    Private Shared _accesser As DateTimeUtils = New DateTimeUtils()

    '現在の時刻
    Private _nowTime As Date

    '初期設定の分割時間
    Private _divMin As Integer = DEFAULT_DIV_TIME

    'イベント定義
    Public Event PopTimerEvent(ByVal sender As Object, ByVal e As DateTimeUtilsEventArgs)



    ' コンストラクタです。(外部からのアクセス不可)
    Private Sub New()
        _nowTime = DateTime.Now
        '作成時はタイマーとしての機能はOFF
        Me.Enabled = False
        Me.Interval = DEFAULT_INTERVAL
    End Sub

    ''' <summary>
    ''' インスタンスを取得します。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getInstance() As DateTimeUtils
        Return _accesser
    End Function


    ''' <summary>
    ''' 大体の時刻を求める。
    ''' </summary>
    ''' <param name="d"></param>
    ''' <param name="min"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRoundDateTime(Optional ByVal min As Integer = 0, Optional ByVal d As DateTime = Nothing) As DateTime
        If min < 1 Then
            If d <> Nothing Then
                If (_nowTime.Minute Mod _divMin) * 2 > _divMin Then
                    getRoundDateTime = d.AddMinutes(min - (d.Minute Mod _divMin))
                Else
                    getRoundDateTime = d.AddMinutes(-(d.Minute Mod _divMin))
                End If
            Else
                If (_nowTime.Minute Mod _divMin) * 2 > _divMin Then
                    getRoundDateTime = _nowTime.AddMinutes(min - (_nowTime.Minute Mod _divMin))
                Else
                    getRoundDateTime = _nowTime.AddMinutes(-(_nowTime.Minute Mod _divMin))
                End If
            End If
        Else
            If d <> Nothing Then
                If (_nowTime.Minute Mod min) * 2 > min Then
                    getRoundDateTime = d.AddMinutes(min - (d.Minute Mod min))
                Else
                    getRoundDateTime = d.AddMinutes(-(d.Minute Mod min))
                End If
            Else
                If (_nowTime.Minute Mod min) * 2 > min Then
                    getRoundDateTime = _nowTime.AddMinutes(min - (_nowTime.Minute Mod min))
                Else
                    getRoundDateTime = _nowTime.AddMinutes(-(_nowTime.Minute Mod min))
                End If
            End If
        End If
    End Function


    ''' <summary>
    ''' 内部保持現在時刻を更新
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub updateTime(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Tick
        _nowTime = DateTime.Now
        RaiseEvent PopTimerEvent(Me, New DateTimeUtilsEventArgs(getRoundDateTime(_divMin)))
    End Sub



    ''' <summary>
    ''' 分割時間のプロパティ。単位：分。1以下の場合、自動的に1となる。
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DivideInMinutes() As Integer
        Get
            Return _divMin
        End Get
        Set(ByVal value As Integer)
            If value < 1 Then
                _divMin = 1
            Else
                _divMin = value
            End If
        End Set
    End Property


End Class

Public Class DateTimeUtilsEventArgs : Inherits EventArgs

    'TimerEventで発生する時刻
    Private _dateTime As DateTime

    ''' <summary>
    ''' コンストラクタ。
    ''' </summary>
    ''' <param name="time"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal time As DateTime)
        _dateTime = time
    End Sub



    ''' <summary>プログレスバーの進捗状況を取得します。</summary>
    Public Property time() As Date
        Get
            Return _dateTime
        End Get
        Set(ByVal value As Date)
            _dateTime = value
        End Set
    End Property

End Class
