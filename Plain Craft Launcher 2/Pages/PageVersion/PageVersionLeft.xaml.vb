Public Class PageVersionLeft
    Implements IRefreshable

    ''' <summary>
    ''' 当前显示设置的 MC 版本。
    ''' </summary>
    Public Shared Version As McVersion = Nothing

    Public Sub RefreshModDisabled() Handles Me.Loaded
        If Version IsNot Nothing AndAlso Version.Modable Then
            ItemMod.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionMod"), Visibility.Collapsed, Visibility.Visible)
            ItemModDisabled.Visibility = Visibility.Collapsed
        Else
            ItemMod.Visibility = Visibility.Collapsed
            ItemModDisabled.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionMod"), Visibility.Collapsed, Visibility.Visible)
        End If
        '功能隐藏
        If Not PageSetupUI.HiddenForceShow Then
            Dim DisableCount As Integer = 0
            If Setup.Get("UiHiddenVersionSave") Then DisableCount += 1
            If Setup.Get("UiHiddenVersionScreenshot") Then DisableCount += 1
            If Setup.Get("UiHiddenVersionMod") Then DisableCount += 1
            If Setup.Get("UiHiddenVersionResourcePack") Then DisableCount += 1
            If Setup.Get("UiHiddenVersionShader") Then DisableCount += 1
            If Setup.Get("UiHiddenVersionSchematic") Then DisableCount += 1
            If DisableCount = 6 Then
                TextResource.Visibility = Visibility.Collapsed
            Else
                TextResource.Visibility = Visibility.Visible
            End If
        Else
            TextResource.Visibility = Visibility.Visible
        End If
        ItemInstall.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionEdit"), Visibility.Collapsed, Visibility.Visible)
        ItemExport.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionExport"), Visibility.Collapsed, Visibility.Visible)
        ItemWorld.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionSave"), Visibility.Collapsed, Visibility.Visible)
        ItemScreenshot.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionScreenshot"), Visibility.Collapsed, Visibility.Visible)
        ItemResourcePack.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionResourcePack"), Visibility.Collapsed, Visibility.Visible)
        ItemShader.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionShader"), Visibility.Collapsed, Visibility.Visible)
        ItemSchematic.Visibility = If(Not PageSetupUI.HiddenForceShow AndAlso Setup.Get("UiHiddenVersionSchematic"), Visibility.Collapsed, Visibility.Visible)
    End Sub

#Region "页面切换"

    ''' <summary>
    ''' 当前页面的编号。从 0 开始计算。
    ''' </summary>
    Public PageID As FormMain.PageSubType = FormMain.PageSubType.Default

    ''' <summary>
    ''' 勾选事件改变页面。
    ''' </summary>
    Private Sub PageCheck(sender As MyListItem, e As RouteEventArgs) Handles ItemOverall.Check, ItemMod.Check, ItemModDisabled.Check, ItemSetup.Check, ItemScreenshot.Check, ItemWorld.Check, ItemResourcePack.Check, ItemShader.Check, ItemSchematic.Check, ItemInstall.Check, ItemExport.Check
        '尚未初始化控件属性时，sender.Tag 为 Nothing，会导致切换到页面 0
        '若使用 IsLoaded，则会导致模拟点击不被执行（模拟点击切换页面时，控件的 IsLoaded 为 False）
        If sender.Tag IsNot Nothing Then PageChange(Val(sender.Tag))
    End Sub

    Public Function PageGet(Optional ID As FormMain.PageSubType = -1)
        If ID = -1 Then ID = PageID
        Select Case ID
            Case FormMain.PageSubType.VersionOverall
                If FrmVersionOverall Is Nothing Then FrmVersionOverall = New PageVersionOverall
                Return FrmVersionOverall
            Case FormMain.PageSubType.VersionMod
                If FrmVersionMod Is Nothing Then FrmVersionMod = New PageVersionCompResource(CompType.Mod)
                Return FrmVersionMod
            Case FormMain.PageSubType.VersionModDisabled
                If FrmVersionModDisabled Is Nothing Then FrmVersionModDisabled = New PageVersionModDisabled
                Return FrmVersionModDisabled
            Case FormMain.PageSubType.VersionSetup
                If IsNothing(FrmVersionSetup) Then FrmVersionSetup = New PageVersionSetup
                Return FrmVersionSetup
            Case FormMain.PageSubType.VersionWorld
                If FrmVersionSaves Is Nothing Then FrmVersionSaves = New PageVersionSaves
                Return FrmVersionSaves
            Case FormMain.PageSubType.VersionScreenshot
                If FrmVersionScreenshot Is Nothing Then FrmVersionScreenshot = New PageVersionScreenshot
                Return FrmVersionScreenshot
            Case FormMain.PageSubType.VersionResourcePack
                If FrmVersionResourcePack Is Nothing Then FrmVersionResourcePack = New PageVersionCompResource(CompType.ResourcePack)
                Return FrmVersionResourcePack
            Case FormMain.PageSubType.VersionShader
                If FrmVersionShader Is Nothing Then FrmVersionShader = New PageVersionCompResource(CompType.Shader)
                Return FrmVersionShader
            Case FormMain.PageSubType.VersionSchematic
                If FrmVersionSchematic Is Nothing Then FrmVersionSchematic = New PageVersionCompResource(CompType.Schematic)
                Return FrmVersionSchematic
            Case FormMain.PageSubType.VersionInstall
                If FrmVersionInstall Is Nothing Then FrmVersionInstall = New PageVersionInstall
                Return FrmVersionInstall
            Case FormMain.PageSubType.VersionExport
                If FrmVersionExport Is Nothing Then FrmVersionExport = New PageVersionExport
                Return FrmVersionExport
            Case Else
                Throw New Exception("未知的版本设置子页面种类：" & ID)
        End Select
    End Function

    ''' <summary>
    ''' 切换现有页面。
    ''' </summary>
    Public Sub PageChange(ID As FormMain.PageSubType)
        If PageID = ID Then Return
        AniControlEnabled += 1
        Try
            PageChangeRun(PageGet(ID))
            PageID = ID
        Catch ex As Exception
            Log(ex, "切换分页面失败（ID " & ID & "）", LogLevel.Feedback)
        Finally
            AniControlEnabled -= 1
        End Try
    End Sub
    Private Shared Sub PageChangeRun(Target As MyPageRight)
        AniStop("FrmMain PageChangeRight") '停止主页面的右页面切换动画，防止它与本动画一起触发多次 PageOnEnter
        If Target.Parent IsNot Nothing Then Target.SetValue(ContentPresenter.ContentProperty, Nothing)
        FrmMain.PageRight = Target
        CType(FrmMain.PanMainRight.Child, MyPageRight).PageOnExit()
        AniStart({
            AaCode(Sub()
                       CType(FrmMain.PanMainRight.Child, MyPageRight).PageOnForceExit()
            FrmMain.PanMainRight.Child = FrmMain.PageRight
                       FrmMain.PageRight.Opacity = 0
                   End Sub, 130),
            AaCode(Sub()
                       '延迟触发页面通用动画，以使得在 Loaded 事件中加载的控件得以处理
                       FrmMain.PageRight.Opacity = 1
                       FrmMain.PageRight.PageOnEnter()
                   End Sub, 30, True)
        }, "PageLeft PageChange")
    End Sub

#End Region

    Public Sub Refresh(sender As Object, e As EventArgs) '由边栏按钮匿名调用
        Refresh(Val(sender.Tag))
    End Sub
    Public Sub Refresh() Implements IRefreshable.Refresh
        Refresh(FrmMain.PageCurrentSub)
    End Sub
    Public Sub Refresh(SubType As FormMain.PageSubType)
        Select Case SubType
            Case FormMain.PageSubType.VersionMod
                PageVersionCompResource.Refresh(CompType.Mod)
            Case FormMain.PageSubType.VersionScreenshot
                PageVersionScreenshot.Refresh()
            Case FormMain.PageSubType.VersionWorld
                PageVersionSaves.Refresh()
            Case FormMain.PageSubType.VersionResourcePack
                PageVersionCompResource.Refresh(CompType.ResourcePack)
            Case FormMain.PageSubType.VersionShader
                PageVersionCompResource.Refresh(CompType.Shader)
            Case FormMain.PageSubType.VersionSchematic
                PageVersionCompResource.Refresh(CompType.Schematic)
            Case FormMain.PageSubType.VersionInstall
                DlClientListLoader.Start(IsForceRestart:=True)
                DlOptiFineListLoader.Start(IsForceRestart:=True)
                DlForgeListLoader.Start(IsForceRestart:=True)
                DlNeoForgeListLoader.Start(IsForceRestart:=True)
                DlLiteLoaderListLoader.Start(IsForceRestart:=True)
                DlFabricListLoader.Start(IsForceRestart:=True)
                DlFabricApiLoader.Start(IsForceRestart:=True)
                DlQuiltListLoader.Start(IsForceRestart:=True)
                DlQSLLoader.Start(IsForceRestart:=True)
                DlOptiFabricLoader.Start(IsForceRestart:=True)
                DlLabyModListLoader.Start(IsForceRestart:=True)
                ItemInstall.Checked = True
                FrmVersionInstall.GetCurrentInfo()
            Case FormMain.PageSubType.VersionExport
                If FrmVersionExport IsNot Nothing Then FrmVersionExport.RefreshAll()
                ItemExport.Checked = True
        End Select
    End Sub

    Public Sub Reset(sender As Object, e As EventArgs)
        If MyMsgBox(GetLang("LangPageVersionLeftSettingDialogIndependentSetContent"), GetLang("LangPageVersionLeftSettingDialogIndependentSetTitle"),, GetLang("LangDialogBtnCancel"), IsWarn:=True) = 1 Then
            If IsNothing(FrmVersionSetup) Then FrmVersionSetup = New PageVersionSetup
            FrmVersionSetup.Reset()
            ItemSetup.Checked = True
        End If
    End Sub

End Class
