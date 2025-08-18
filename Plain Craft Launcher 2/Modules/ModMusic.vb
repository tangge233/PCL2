Imports Windows.Media
Imports Windows.Storage
Imports Windows.Storage.Streams
Public Module ModMusic

#Region "播放列表"

    ''' <summary>
    ''' 接下来要播放的音乐文件路径。未初始化时为 Nothing。
    ''' </summary>
    Public MusicWaitingList As List(Of String) = Nothing
    ''' <summary>
    ''' 全部音乐文件路径。未初始化时为 Nothing。
    ''' </summary>
    Public MusicAllList As List(Of String) = Nothing
    ''' <summary>
    ''' 初始化音乐播放列表。
    ''' </summary>
    ''' <param name="ForceReload">强制全部重新载入列表。</param>
    ''' <param name="PreventFirst">在重载列表时避免让某项成为第一项。</param>
    Private Sub MusicListInit(ForceReload As Boolean, Optional PreventFirst As String = Nothing)
        If ForceReload Then MusicAllList = Nothing
        Try
            '初始化全部可用音乐列表
            If MusicAllList Is Nothing Then
                MusicAllList = New List(Of String)
                Directory.CreateDirectory(Path & "PCL\Musics\")
                For Each File In EnumerateFiles(Path & "PCL\Musics\")
                    '文件夹可能会被加入 .ini 文件夹配置文件、一些乱七八糟的 .jpg 文件啥的
                    Dim Ext As String = File.Extension.ToLower
                    If {".ini", ".jpg", ".txt", ".cfg", ".lrc", ".db", ".png"}.Contains(Ext) Then Continue For
                    MusicAllList.Add(File.FullName)
                Next
            End If
            '打乱顺序播放
            MusicWaitingList = If(Setup.Get("UiMusicRandom"), Shuffle(New List(Of String)(MusicAllList)), New List(Of String)(MusicAllList))
            If PreventFirst IsNot Nothing AndAlso MusicWaitingList.FirstOrDefault = PreventFirst Then
                '若需要避免成为第一项的为第一项，则将它放在最后
                MusicWaitingList.RemoveAt(0)
                MusicWaitingList.Add(PreventFirst)
            End If
        Catch ex As Exception
            Log(ex, GetLang("LangModMusicExceptionInitFail"), LogLevel.Feedback)
        End Try
    End Sub
    ''' <summary>
    ''' 获取下一首播放的音乐路径并将其从列表中移除。
    ''' 如果没有，可能会返回 Nothing。
    ''' </summary>
    Private Function DequeueNextMusicAddress() As String
        '初始化，确保存在音乐
        If MusicAllList Is Nothing OrElse Not MusicAllList.Any() OrElse Not MusicWaitingList.Any() Then MusicListInit(False)
        '出列下一个音乐，如果出列结束则生成新列表
        If MusicWaitingList.Any() Then
            DequeueNextMusicAddress = MusicWaitingList(0)
            MusicWaitingList.RemoveAt(0)
        Else
            DequeueNextMusicAddress = Nothing
        End If
        If Not MusicWaitingList.Any() Then MusicListInit(False, DequeueNextMusicAddress)
    End Function

#End Region

#Region "UI 控制"

    ''' <summary>
    ''' 刷新背景音乐按钮 UI 与设置页 UI。
    ''' </summary>
    Private Sub MusicRefreshUI()
        RunInUi(
        Sub()
            Try

                If Not MusicAllList.Any() Then
                    '无背景音乐
                    FrmMain.BtnExtraMusic.Show = False
                Else
                    '有背景音乐
                    FrmMain.BtnExtraMusic.Show = True
                    Dim ToolTipText As String
                    If MusicState = MusicStates.Pause Then
                        FrmMain.BtnExtraMusic.Logo = Logo.IconPlay
                        FrmMain.BtnExtraMusic.LogoScale = 0.8
                        ToolTipText = GetLang("LangModMusicPaused", GetFileNameWithoutExtentionFromPath(MusicCurrent))
                        If MusicAllList.Count > 1 Then
                            ToolTipText += vbCrLf & GetLang("LangModMusicStopClickTipA")
                        Else
                            ToolTipText += vbCrLf & GetLang("LangModMusicStopClickTipB")
                        End If
                    Else
                        FrmMain.BtnExtraMusic.Logo = Logo.IconMusic
                        FrmMain.BtnExtraMusic.LogoScale = 1
                        ToolTipText = GetLang("LangModMusicPlaying", GetFileNameWithoutExtentionFromPath(MusicCurrent))
                        If MusicAllList.Count > 1 Then
                            ToolTipText += vbCrLf & GetLang("LangModMusicStartClickTipA")
                        Else
                            ToolTipText += vbCrLf & GetLang("LangModMusicStartClickTipB")
                        End If
                    End If
                    FrmMain.BtnExtraMusic.ToolTip = ToolTipText
                    ToolTipService.SetVerticalOffset(FrmMain.BtnExtraMusic, If(ToolTipText.Contains(vbLf), 10, 16))
                End If
                If FrmSetupUI IsNot Nothing Then FrmSetupUI.MusicRefreshUI()

            Catch ex As Exception
                Log(ex, GetLang("LangModMusicExceptionUIRefreshFail"), LogLevel.Feedback)
            End Try
        End Sub)
        Dim SMTCUpdater = _smtc?.DisplayUpdater
        If SMTCUpdater IsNot Nothing Then
            Try
                Using TagInfo = TagLib.File.Create(MusicCurrent)
                    Dim Tags = TagInfo.Tag
                    ' 标题
                    SMTCUpdater.MusicProperties.Title = If(Not String.IsNullOrEmpty(Tags.Title),
                                                  Tags.Title,
                                                  GetFileNameWithoutExtentionFromPath(MusicCurrent))
                    ' 作者
                    Dim firstArtist As String = If(Tags.Performers.Any(),
                                          Tags.Performers(0),
                                          If(Not String.IsNullOrEmpty(Tags.FirstPerformer),
                                             Tags.FirstPerformer,
                                             ""))
                    SMTCUpdater.MusicProperties.Artist = firstArtist
                    ' 专辑作者
                    Dim albumArtist As String = If(Tags.AlbumArtists.Any(),
                                           Tags.AlbumArtists(0),
                                           If(Not String.IsNullOrEmpty(Tags.FirstAlbumArtist),
                                              Tags.FirstAlbumArtist,
                                              ""))
                    SMTCUpdater.MusicProperties.AlbumArtist = albumArtist
                    SMTCUpdater.MusicProperties.AlbumTitle = If(Tags.Album, "")
                    ' 处理专辑封面（使用嵌入的图片数据）
                    If Tags.Pictures.Any() Then
                        Dim firstPicture = Tags.Pictures(0)
                        Dim targetPath = $"{PathTemp}Cache\albumcover.jpg"
                        File.WriteAllBytes(targetPath, firstPicture.Data.Data)
                        Dim tempFile = StorageFile.GetFileFromPathAsync(targetPath).GetResults()
                        ' 写入图片数据
                        SMTCUpdater.Thumbnail = RandomAccessStreamReference.CreateFromFile(tempFile)
                    Else
                        SMTCUpdater.Thumbnail = Nothing
                    End If
                End Using
            Catch ex As Exception
                SMTCUpdater.MusicProperties.Title = GetFileNameWithoutExtentionFromPath(MusicCurrent)
                SMTCUpdater.MusicProperties.Artist = ""
                SMTCUpdater.MusicProperties.AlbumArtist = ""
                SMTCUpdater.MusicProperties.AlbumTitle = ""
                SMTCUpdater.Thumbnail = Nothing
                Log(ex, "[SMTC] 读取音乐详细信息出错")
            End Try
            SMTCUpdater.Update()
        End If
    End Sub

    ''' <summary>
    ''' 让音乐在暂停、播放间切换，并显示提示文本。
    ''' </summary>
    Public Sub MusicControlPause()
        If MusicNAudio Is Nothing Then
            Hint(GetLang("LangModMusicNotPlayedYet"), HintType.Critical)
        Else
            Select Case MusicState
                Case MusicStates.Pause
                    MusicResume()
                Case MusicStates.Play
                    MusicPause()
                Case Else
                    Log("[Music] 音乐目前为停止状态，已强制尝试开始播放", LogLevel.Debug)
                    MusicRefreshPlay(False)
            End Select
        End If
    End Sub

    ''' <summary>
    ''' 播放下一曲，并显示提示文本。
    ''' </summary>
    Public Sub MusicControlNext()
        If MusicAllList.Count = 1 Then
            MusicStartPlay(MusicCurrent)
            Hint(GetLang("LangModMusicReplay", GetFileNameFromPath(MusicCurrent)), HintType.Finish)
        Else
            Dim Address As String = DequeueNextMusicAddress()
            If Address Is Nothing Then
                Hint(GetLang("LangModMusicNoMusic"), HintType.Critical)
            Else
                MusicStartPlay(Address)
                Hint(GetLang("LangModMusicPlaying", GetFileNameFromPath(Address)), HintType.Finish)
            End If
        End If
        MusicRefreshUI()
    End Sub

#End Region

#Region "主状态控制"

    ''' <summary>
    ''' 获取当前的音乐播放状态。
    ''' </summary>
    Public ReadOnly Property MusicState As MusicStates
        Get
            If MusicNAudio Is Nothing Then Return MusicStates.Stop
            Select Case MusicNAudio.PlaybackState
                Case 0 'NAudio.Wave.PlaybackState.Stopped
                    Return MusicStates.Stop
                Case 2 'NAudio.Wave.PlaybackState.Paused
                    Return MusicStates.Pause
                Case Else
                    Return MusicStates.Play
            End Select
        End Get
    End Property
    Public Enum MusicStates
        [Stop]
        Play
        Pause
    End Enum

    '重载与开始

    ''' <summary>
    ''' 重载播放列表并尝试开始播放背景音乐。
    ''' </summary>
    ''' <param name="ShowHint">是否显示刷新提示。</param>
    Public Sub MusicRefreshPlay(ShowHint As Boolean, Optional IsFirstLoad As Boolean = False)
        Try

            MusicListInit(True)
            If Not MusicAllList.Any() Then
                If MusicNAudio Is Nothing Then
                    If ShowHint Then Hint(GetLang("LangModMusicNoMusic"), HintType.Critical)
                Else
                    MusicNAudio = Nothing
                    If ShowHint Then Hint(GetLang("LangModMusicMusicCleared"), HintType.Finish)
                End If
            Else
                Dim Address As String = DequeueNextMusicAddress()
                If Address Is Nothing Then
                    If ShowHint Then Hint(GetLang("LangModMusicNoMusic"), HintType.Critical)
                Else
                    Try
                        MusicStartPlay(Address, IsFirstLoad)
                        If ShowHint Then Hint(GetLang("LangModMusicMusicRefreshed", GetFileNameFromPath(Address)), HintType.Finish, False)
                    Catch
                    End Try
                End If
            End If
            MusicRefreshUI()

        Catch ex As Exception
            Log(ex, GetLang("LangModMusicExceptionMusicRefreshFail"), LogLevel.Feedback)
        End Try
    End Sub
    ''' <summary>
    ''' 开始播放音乐。
    ''' </summary>
    Private Sub MusicStartPlay(Address As String, Optional IsFirstLoad As Boolean = False)
        If Address Is Nothing Then Return
        If _smtc Is Nothing OrElse Not _smtc.IsEnabled Then EnableSMTCSupport()
        Log("[Music] 播放开始：" & Address)
        MusicCurrent = Address
        UpdateSMTCInfo()
        RunInNewThread(Sub() MusicLoop(IsFirstLoad), "Music", ThreadPriority.BelowNormal)
    End Sub

    '播放与暂停

    ''' <summary>
    ''' 暂停音乐播放，返回是否成功切换了状态。
    ''' </summary>
    Public Function MusicPause() As Boolean
        If MusicState = MusicStates.Play Then
            RunInThread(
            Sub()
                Log("[Music] 已暂停播放")
                MusicNAudio?.Pause()
                SetSMTCStatus()
                MusicRefreshUI()
            End Sub)
            Return True
        Else
            Log($"[Music] 无需暂停播放，当前状态为 {MusicState}")
            Return False
        End If
    End Function
    ''' <summary>
    ''' 继续音乐播放，返回是否成功切换了状态。
    ''' </summary>
    Public Function MusicResume() As Boolean
        If MusicState = MusicStates.Play OrElse Not MusicAllList.Any() Then
            Log($"[Music] 无需继续播放，当前状态为 {MusicState}")
            Return False
        Else
            RunInThread(
            Sub()
                Log("[Music] 已恢复播放")
                Try
                    MusicNAudio?.Play()
                Catch 'https://github.com/Meloong-Git/PCL/pull/5415#issuecomment-2751135223
                    MusicNAudio?.Stop()
                    MusicNAudio?.Play()
                End Try
                SetSMTCStatus()
                MusicRefreshUI()
            End Sub)
            Return True
        End If
    End Function

#End Region

#Region "SMTC 控件"
    Private ReadOnly _player = If(Environment.OSVersion.Version.ToString().Substring(0, 4) = "10.0", New Playback.MediaPlayer, Nothing)
    Private _smtc As SystemMediaTransportControls = Nothing

    ''' <summary>
    ''' 启用 SMTC 支持
    ''' </summary>
    Public Sub EnableSMTCSupport()
        If Not Environment.OSVersion.Version.ToString().Substring(0, 4) = "10.0" Then
            Log("[SMTC] 当前系统不支持 SMTC 控件，不进行 SMTC 控件初始化")
            Exit Sub
        End If
        If Not Setup.Get("UiMusicSMTC") Then
            Log("[SMTC] 用户已关闭 SMTC 支持，不进行初始化")
            Exit Sub
        End If
        If _smtc IsNot Nothing AndAlso _smtc.IsEnabled Then
            Log("[SMTC] 已存在 SMTC 信息源，无需重复初始化")
            Exit Sub
        End If
        Log("[SMTC] 初始化 SMTC 支持")
        '初始化
        _player.CommandManager.IsEnabled = False
        _smtc = _player.SystemMediaTransportControls
        _smtc.IsEnabled = True
        Dim updater = _smtc.DisplayUpdater
        updater.AppMediaId = "Plain Craft Launcher Community Edition" '媒体来源信息
        updater.Type = MediaPlaybackType.Music '指定媒体类型
        updater.Update()

        '设置可交互性
        _smtc.IsPlayEnabled = True
        _smtc.IsPauseEnabled = True
        _smtc.IsNextEnabled = True
        _smtc.IsPreviousEnabled = False '暂时没有上一首

        '绑定事件处理
        AddHandler _smtc.ButtonPressed, AddressOf _smtc_ButtonPressed
    End Sub

    ''' <summary>
    ''' 关闭 SMTC 支持
    ''' </summary>
    Public Sub DisableSMTCSupport()
        If _smtc IsNot Nothing AndAlso _smtc.IsEnabled Then
            Log("[SMTC] 移除 SMTC 信息源")
            _smtc.IsEnabled = False
            _smtc = Nothing
        Else
            Log("[SMTC] 未添加 SMTC 信息，无需移除")
        End If
    End Sub

    ''' <summary>
    ''' 更新 SMTC 信息
    ''' </summary>
    Public Async Sub UpdateSMTCInfo()
        If _smtc Is Nothing Then Exit Sub
        Log($"[SMTC] 更新 SMTC 媒体信息，文件路径: {MusicCurrent}")
        Dim Updater = _smtc.DisplayUpdater

        Try
            Dim File = TagLib.File.Create(MusicCurrent)

            Updater.AppMediaId = "Plain Craft Launcher 2 CE" '媒体来源信息
            Updater.Type = MediaPlaybackType.Music '指定媒体类型

            Dim Artist As String = File.Tag.FirstPerformer
            Dim AlbumArtist As String = File.Tag.FirstAlbumArtist
            Dim AlbumTitle As String = File.Tag.Album
            Dim Title As String = File.Tag.Title
            Dim Thumbnail = File.Tag.Pictures.FirstOrDefault()

            If String.IsNullOrEmpty(Artist) Then
                Artist = ""
            End If
            If String.IsNullOrEmpty(AlbumArtist) Then
                AlbumArtist = Artist
            End If
            If String.IsNullOrEmpty(AlbumTitle) Then
                AlbumTitle = ""
            End If
            If String.IsNullOrEmpty(Title) Then
                Title = ""
            End If

            Updater.MusicProperties.Artist = Artist
            Updater.MusicProperties.AlbumArtist = AlbumArtist
            Updater.MusicProperties.AlbumTitle = AlbumTitle
            Updater.MusicProperties.Title = Title

            If Thumbnail IsNot Nothing Then
                Dim MemStream As InMemoryRandomAccessStream = New InMemoryRandomAccessStream()
                Using Writer As New DataWriter(MemStream)
                    Writer.WriteBytes(Thumbnail.Data.Data)
                    Await Writer.StoreAsync()
                    Writer.DetachStream()
                End Using
                Dim ThumbailStream = RandomAccessStreamReference.CreateFromStream(MemStream)

                Updater.Thumbnail = ThumbailStream
            Else
                Updater.Thumbnail = Nothing
            End If
        Catch
        End Try

        '生效设置
        Updater.Update()
    End Sub

    ''' <summary>
    ''' 设置 SMTC 媒体播放状态
    ''' </summary>
    Public Sub SetSMTCStatus()
        If _smtc Is Nothing Then Exit Sub
        If MusicState = MusicStates.Play Then
            Log("[SMTC] 更新 SMTC 播放状态为：Playing")
            _smtc.PlaybackStatus = MediaPlaybackStatus.Playing
        ElseIf MusicState = MusicStates.Pause Then
            Log("[SMTC] 更新 SMTC 播放状态为：Paused")
            _smtc.PlaybackStatus = MediaPlaybackStatus.Paused
        Else
            Log("[SMTC] 更新 SMTC 播放状态为：Stopped")
            _smtc.PlaybackStatus = MediaPlaybackStatus.Stopped
        End If
    End Sub

    ''' <summary>
    ''' 响应 SMTC 交互
    ''' </summary>
    Public Sub _smtc_ButtonPressed(sender As SystemMediaTransportControls, args As SystemMediaTransportControlsButtonPressedEventArgs)
        If _smtc Is Nothing Then Exit Sub
        Select Case args.Button
            Case SystemMediaTransportControlsButton.Play
                Log("[SMTC] 收到 SMTC 控件事件，切换播放状态为：Playing")
                MusicResume()
                Exit Select
            Case SystemMediaTransportControlsButton.Pause
                Log("[SMTC] 收到 SMTC 控件事件，切换播放状态为：Paused")
                MusicPause()
                Exit Select
            Case SystemMediaTransportControlsButton.Next
                Log("[SMTC] 收到 SMTC 控件事件，切换到下一曲")
                MusicControlNext()
                Exit Select
        End Select
    End Sub

    ''' <summary>
    ''' 更新 SMTC 时间线属性
    ''' </summary>
    ''' <param name="CurrentTime">当前播放进度</param>
    ''' <param name="TotalTime">曲目全长</param>
    Public Sub UpdateSMTCTimeline(CurrentTime As TimeSpan, TotalTime As TimeSpan)
        Dim Properties = New SystemMediaTransportControlsTimelineProperties With {
            .StartTime = TimeSpan.FromSeconds(0),
            .MinSeekTime = TimeSpan.FromSeconds(0),
            .Position = CurrentTime,
            .MaxSeekTime = CurrentTime,
            .EndTime = TotalTime
        }

        _smtc.UpdateTimelineProperties(Properties)
    End Sub

    ''' <summary>
    ''' 以 700 ms 为刷新间隔的 SMTC 时间线更新
    ''' </summary>
    Public Sub SMTCTimelineUpdater(CurrentWave As NAudio.Wave.WaveOutEvent, Reader As NAudio.Wave.WaveStream)
        If _smtc Is Nothing Then Exit Sub
        RunInNewThread(Sub()
                           While CurrentWave.Equals(MusicNAudio) AndAlso CurrentWave.PlaybackState = NAudio.Wave.PlaybackState.Playing AndAlso _smtc IsNot Nothing
                               RunInNewThread(Sub() UpdateSMTCTimeline(Reader.CurrentTime, Reader.TotalTime))
                               Thread.Sleep(700)
                           End While
                           While CurrentWave.Equals(MusicNAudio) AndAlso CurrentWave.PlaybackState = NAudio.Wave.PlaybackState.Paused AndAlso _smtc IsNot Nothing
                               Thread.Sleep(700)
                           End While
                           If Not CurrentWave.Equals(MusicNAudio) Then Exit Sub
                           SMTCTimelineUpdater(CurrentWave, Reader)
                       End Sub, "SMTC Updater")
    End Sub
#End Region

    ''' <summary>
    ''' 当前正在播放的 NAudio.Wave.WaveOutEvent。
    ''' </summary>
    Public MusicNAudio = Nothing
    ''' <summary>
    ''' 当前播放的音乐地址。
    ''' </summary>
    Private MusicCurrent As String = ""
    ''' <summary>
    ''' 在 MusicUuid 不变的前提下，持续播放某地址的音乐，且在播放结束后随机播放下一曲。
    ''' </summary>
    Private Sub MusicLoop(Optional IsFirstLoad As Boolean = False)
        Dim CurrentWave As NAudio.Wave.WaveOutEvent = Nothing
        Dim Reader As NAudio.Wave.WaveStream = Nothing
        Try
            '开始播放
            CurrentWave = New NAudio.Wave.WaveOutEvent()
            MusicNAudio = CurrentWave
            CurrentWave.DeviceNumber = -1
            Reader = New NAudio.Wave.AudioFileReader(MusicCurrent)
            CurrentWave.Init(Reader)
            CurrentWave.Play()
            '第一次打开的暂停
            If IsFirstLoad AndAlso Not Setup.Get("UiMusicAuto") Then
                CurrentWave.Pause()
                EnableSMTCSupport() '启用 SMTC 支持
                UpdateSMTCInfo() '更新 SMTC 媒体信息
            End If
            SetSMTCStatus()
            MusicRefreshUI()
            '停止条件：播放完毕或变化
            Dim PreviousVolume = 0
            SMTCTimelineUpdater(CurrentWave, Reader) '启动 SMTC 时间轴更新
            While CurrentWave.Equals(MusicNAudio) AndAlso Not CurrentWave.PlaybackState = NAudio.Wave.PlaybackState.Stopped
                If Setup.Get("UiMusicVolume") <> PreviousVolume Then
                    '更新音量
                    PreviousVolume = Setup.Get("UiMusicVolume")
                    CurrentWave.Volume = PreviousVolume / 1000
                End If
                '更新进度条
                Dim Percent = Reader.CurrentTime.TotalMilliseconds / Reader.TotalTime.TotalMilliseconds
                RunInUi(Sub() FrmMain.BtnExtraMusic.Progress = Percent)
                '检查 SMTC 状态
                If Setup.Get("UiMusicSMTC") AndAlso _smtc Is Nothing Then
                    EnableSMTCSupport()
                    UpdateSMTCInfo()
                    SetSMTCStatus()
                    SMTCTimelineUpdater(CurrentWave, Reader)
                End If
                If Not Setup.Get("UiMusicSMTC") AndAlso _smtc IsNot Nothing Then DisableSMTCSupport()
                Thread.Sleep(100)
            End While
            '当前音乐已播放结束，继续下一曲
            If CurrentWave.PlaybackState = NAudio.Wave.PlaybackState.Stopped AndAlso MusicAllList.Any Then MusicStartPlay(DequeueNextMusicAddress)
        Catch ex As Exception
            Log(ex, "播放音乐出现内部错误（" & MusicCurrent & "）", LogLevel.Developer)
            If TypeOf ex Is NAudio.MmException AndAlso ex.Message.Contains("AlreadyAllocated") Then
                Hint("你的音频设备正被其他程序占用。请在关闭占用的程序后重启 PCL，才能恢复音乐播放功能！", HintType.Critical)
                Thread.Sleep(1000000000)
            End If
            If TypeOf ex Is NAudio.MmException AndAlso (ex.Message.Contains("NoDriver") OrElse ex.Message.Contains("BadDeviceId")) Then
                Hint(GetLang("LangModMusicDeviceChanged"), HintType.Critical)
                Thread.Sleep(1000000000)
            End If
            If ex.Message.Contains("Got a frame at sample rate") OrElse ex.Message.Contains("does not support changes to") Then
                Hint(GetLang("LangModMusicMusicChanged", GetFileNameFromPath(MusicCurrent)), HintType.Critical)
            ElseIf Not (MusicCurrent.EndsWithF(".wav", True) OrElse MusicCurrent.EndsWithF(".mp3", True) OrElse MusicCurrent.EndsWithF(".flac", True)) OrElse
                ex.Message.Contains("0xC00D36C4") Then '#5096：不支持给定的 URL 的字节流类型。 (异常来自 HRESULT:0xC00D36C4)
                Hint(GetLang("LangModMusicMusicFormatNotSupport", GetFileNameFromPath(MusicCurrent)), HintType.Critical)
            Else
                Log(ex, "播放音乐失败（" & GetFileNameFromPath(MusicCurrent) & "）", LogLevel.Hint)
            End If
            '将播放错误的音乐从列表中移除
            MusicAllList.Remove(MusicCurrent)
            MusicWaitingList.Remove(MusicCurrent)
            MusicRefreshUI()
            '等待 2 秒后继续播放
            Thread.Sleep(2000)
            If TypeOf ex Is FileNotFoundException Then
                MusicRefreshPlay(True, IsFirstLoad)
            Else
                MusicStartPlay(DequeueNextMusicAddress(), IsFirstLoad)
            End If
        Finally
            If CurrentWave IsNot Nothing Then CurrentWave.Dispose()
            If Reader IsNot Nothing Then Reader.Dispose()
            MusicRefreshUI()
        End Try
    End Sub

End Module
