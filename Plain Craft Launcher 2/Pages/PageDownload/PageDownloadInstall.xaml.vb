Public Class PageDownloadInstall

    Private Sub LoaderInit() Handles Me.Initialized
        DisabledPageAnimControls.Add(BtnStart)
        PageLoaderInit(LoadMinecraft, PanLoad, PanAllBack, Nothing, DlClientListLoader, AddressOf LoadMinecraft_OnFinish)
    End Sub

    Private IsLoad As Boolean = False
    Private Sub Init() Handles Me.Loaded
        PanBack.ScrollToHome()
        DlOptiFineListLoader.Start()
        DlLiteLoaderListLoader.Start()
        DlFabricListLoader.Start()
        DlQuiltListLoader.Start()
        DlNeoForgeListLoader.Start()
        DlCleanroomListLoader.Start()
        DlLabyModListLoader.Start()

        '重载预览
        TextSelectName.ValidateRules = New ObjectModel.Collection(Of Validate) From {New ValidateFolderName(PathMcFolder & "versions")}
        TextSelectName.Validate()
        SelectReload()

        '非重复加载部分
        If IsLoad Then Return
        IsLoad = True

        McDownloadForgeRecommendedRefresh()

        LoadOptiFine.State = DlOptiFineListLoader
        LoadLiteLoader.State = DlLiteLoaderListLoader
        LoadFabric.State = DlFabricListLoader
        LoadFabricApi.State = DlFabricApiLoader
        LoadQuilt.State = DlQuiltListLoader
        LoadQSL.State = DlQSLLoader
        LoadNeoForge.State = DlNeoForgeListLoader
        LoadCleanroom.State = DlCleanroomListLoader
        LoadOptiFabric.State = DlOptiFabricLoader
        LoadLabyMod.State = DlLabyModListLoader
    End Sub

#Region "页面切换"

    '页面切换动画
    Public IsInSelectPage As Boolean = False
    Private IsFirstLoaded As Boolean = False
    Private Sub EnterSelectPage()
        If IsInSelectPage Then Return
        IsInSelectPage = True

        PanInner.Margin = New Thickness(25, 10, 25, 40)

        AutoSelectedFabricApi = False
        AutoSelectedQSL = False
        AutoSelectedOptiFabric = False
        IsSelectNameEdited = False
        PanSelect.Visibility = Visibility.Visible
        PanSelect.IsHitTestVisible = True
        PanMinecraft.IsHitTestVisible = False
        PanBack.IsHitTestVisible = False
        PanBack.ScrollToHome()

        DisabledPageAnimControls.Remove(BtnStart)
        BtnStart.Show = True
        CardOptiFine.IsSwaped = True
        CardLiteLoader.IsSwaped = True
        CardForge.IsSwaped = True
        CardNeoForge.IsSwaped = True
        CardCleanroom.IsSwaped = True
        CardFabric.IsSwaped = True
        CardFabricApi.IsSwaped = True
        CardQuilt.IsSwaped = True
        CardQSL.IsSwaped = True
        CardOptiFabric.IsSwaped = True
        CardLabyMod.IsSwaped = True

        If Not Setup.Get("HintInstallBack") Then
            Setup.Set("HintInstallBack", True)
            Hint(GetLang("LangDownloadInstallHintBackTip"))
        End If

        '如果在选择页面按了刷新键，选择页的东西可能会由于动画被隐藏，但不会由于加载结束而再次显示，因此这里需要手动恢复
        For Each Control In GetAllAnimControls(PanSelect)
            Control.Opacity = 1
            If Control.RenderTransform Is Nothing OrElse TypeOf Control.RenderTransform Is TranslateTransform Then
                Control.RenderTransform = New TranslateTransform
            End If
        Next

        '启动 Forge 加载
        If SelectedMinecraftId.StartsWith("1.") Then
            Dim ForgeLoader = New LoaderTask(Of String, List(Of DlForgeVersionEntry))("DlForgeVersion " & SelectedMinecraftId, AddressOf DlForgeVersionMain)
            LoadForge.State = ForgeLoader
            ForgeLoader.Start(SelectedMinecraftId)
        End If

        '启动 Fabric API、QSL、OptiFabric、LabyMod 加载
        DlFabricApiLoader.Start()
        DlQSLLoader.Start()
        DlOptiFabricLoader.Start()
        DlLabyModListLoader.Start()

        AniStart({
            AaOpacity(PanMinecraft, -PanMinecraft.Opacity, 70, 10),
            AaTranslateX(PanMinecraft, -50 - CType(PanMinecraft.RenderTransform, TranslateTransform).X, 90, 10),
            AaCode(
            Sub()
                PanBack.ScrollToHome()
                TextSelectName.Validate()
                OptiFine_Loaded()
                LiteLoader_Loaded()
                Forge_Loaded()
                NeoForge_Loaded()
                Cleanroom_Loaded()
                Fabric_Loaded()
                FabricApi_Loaded()
                Quilt_Loaded()
                QSL_Loaded()
                OptiFabric_Loaded()
                LabyMod_Loaded()
                SelectReload()
                PanMinecraft.Visibility = Visibility.Collapsed
            End Sub, After:=True),
            AaOpacity(PanSelect, 1 - PanSelect.Opacity, 70, 100),
            AaTranslateX(PanSelect, -CType(PanSelect.RenderTransform, TranslateTransform).X, 160, 100, Ease:=New AniEaseOutFluent(AniEasePower.ExtraStrong)),
            AaCode(
            Sub()
                PanBack.IsHitTestVisible = True
                '初始化 Binding
                If IsFirstLoaded Then Return
                IsFirstLoaded = True
                BtnOptiFineClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardOptiFine.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnLiteLoaderClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardLiteLoader.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnForgeClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardForge.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnNeoForgeClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardNeoForge.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnCleanroomClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardCleanroom.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnFabricClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardFabric.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnFabricApiClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardFabricApi.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnQuiltClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardQuilt.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnQSLClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardQSL.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnLabyModClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardLabyMod.MainTextBlock, .Mode = BindingMode.OneWay})
                BtnOptiFabricClearInner.SetBinding(Shapes.Path.FillProperty, New Binding("Foreground") With {.Source = CardOptiFabric.MainTextBlock, .Mode = BindingMode.OneWay})
            End Sub,, True)
        }, "FrmDownloadInstall SelectPageSwitch", True)
    End Sub
    Public Sub ExitSelectPage() Handles BtnBack.Click
        If Not IsInSelectPage Then Return
        IsInSelectPage = False

        PanInner.Margin = New Thickness(25, 10, 25, 25)

        DisabledPageAnimControls.Add(BtnStart)
        BtnStart.Show = False
        SelectClear() '清除已选择项
        PanMinecraft.Visibility = Visibility.Visible
        PanSelect.IsHitTestVisible = False
        PanMinecraft.IsHitTestVisible = True
        PanBack.IsHitTestVisible = False
        PanBack.ScrollToHome()

        AniStart({
            AaOpacity(PanSelect, -PanSelect.Opacity, 70, 10),
            AaTranslateX(PanSelect, 50 - CType(PanSelect.RenderTransform, TranslateTransform).X, 90, 10),
            AaCode(Sub() PanBack.ScrollToHome(), After:=True),
            AaOpacity(PanMinecraft, 1 - PanMinecraft.Opacity, 70, 100),
            AaTranslateX(PanMinecraft, -CType(PanMinecraft.RenderTransform, TranslateTransform).X, 160, 100, Ease:=New AniEaseOutFluent(AniEasePower.ExtraStrong)),
            AaCode(
            Sub()
                PanSelect.Visibility = Visibility.Collapsed
                PanBack.IsHitTestVisible = True
            End Sub,, True)
        }, "FrmDownloadInstall SelectPageSwitch")
    End Sub
    Public Sub MinecraftSelected(sender As MyListItem, e As MouseButtonEventArgs)
        SelectedMinecraftId = sender.Title
        SelectedMinecraftJsonUrl = sender.Tag("url").ToString
        SelectedMinecraftIcon = sender.Logo
        EnterSelectPage()
    End Sub

#End Region

#Region "选择"

    'Minecraft
    Private SelectedMinecraftId As String
    Private SelectedMinecraftJsonUrl As String
    Private SelectedMinecraftIcon As String

    'OptiFine
    Private SelectedOptiFine As DlOptiFineListEntry = Nothing
    Private Sub SetOptiFineInfoShow(IsShow As String)
        If PanOptiFineInfo.Tag = IsShow Then Return
        PanOptiFineInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanOptiFineInfo, -CType(PanOptiFineInfo.RenderTransform, TranslateTransform).Y, 200, 100, Ease:=New AniEaseOutFluent),
                AaOpacity(PanOptiFineInfo, 1 - PanOptiFineInfo.Opacity, 80, 90)
            }, "SetOptiFineInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanOptiFineInfo, 6 - CType(PanOptiFineInfo.RenderTransform, TranslateTransform).Y, 120),
                AaOpacity(PanOptiFineInfo, -PanOptiFineInfo.Opacity, 50)
            }, "SetOptiFineInfoShow")
        End If
    End Sub

    ''' <summary>
    ''' 选定的 Mod Loader 名称，内容应为 Forge / NeoForge / Fabric / Quilt / Cleanroom / LabyMod
    ''' </summary>
    Private SelectedLoaderName As String = Nothing

    ''' <summary>
    ''' 选定的 Mod Loader API 名称，内容应为 Fabric API 或 QFAPI / QSL
    ''' </summary>
    Private SelectedAPIName As String = Nothing

    'LiteLoader
    Private SelectedLiteLoader As DlLiteLoaderListEntry = Nothing
    Private Sub SetLiteLoaderInfoShow(IsShow As String)
        If PanLiteLoaderInfo.Tag = IsShow Then Return
        PanLiteLoaderInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanLiteLoaderInfo, -CType(PanLiteLoaderInfo.RenderTransform, TranslateTransform).Y, 200, 100, Ease:=New AniEaseOutFluent),
                AaOpacity(PanLiteLoaderInfo, 1 - PanLiteLoaderInfo.Opacity, 80, 90)
            }, "SetLiteLoaderInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanLiteLoaderInfo, 6 - CType(PanLiteLoaderInfo.RenderTransform, TranslateTransform).Y, 120),
                AaOpacity(PanLiteLoaderInfo, -PanLiteLoaderInfo.Opacity, 50)
            }, "SetLiteLoaderInfoShow")
        End If
    End Sub

    'Forge
    Private SelectedForge As DlForgeVersionEntry = Nothing
    Private Sub SetForgeInfoShow(IsShow As String)
        If PanForgeInfo.Tag = IsShow Then Return
        PanForgeInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanForgeInfo, -CType(PanForgeInfo.RenderTransform, TranslateTransform).Y, 200, 100, Ease:=New AniEaseOutFluent),
                AaOpacity(PanForgeInfo, 1 - PanForgeInfo.Opacity, 80, 90)
            }, "SetForgeInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanForgeInfo, 6 - CType(PanForgeInfo.RenderTransform, TranslateTransform).Y, 120),
                AaOpacity(PanForgeInfo, -PanForgeInfo.Opacity, 50)
            }, "SetForgeInfoShow")
        End If
    End Sub

    'Cleanroom
    Private SelectedCleanroom As DlCleanroomListEntry = Nothing
    Private Sub SetCleanroomInfoShow(IsShow As String)
        If PanCleanroomInfo.Tag = IsShow Then Exit Sub
        PanCleanroomInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanCleanroomInfo, -CType(PanCleanroomInfo.RenderTransform, TranslateTransform).Y, 270, 100, Ease:=New AniEaseOutBack),
                AaOpacity(PanCleanroomInfo, 1 - PanCleanroomInfo.Opacity, 100, 90)
            }, "SetCleanroomInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanCleanroomInfo, 6 - CType(PanCleanroomInfo.RenderTransform, TranslateTransform).Y, 200),
                AaOpacity(PanCleanroomInfo, -PanCleanroomInfo.Opacity, 100)
            }, "SetCleanroomInfoShow")
        End If
    End Sub

    'NeoForge
    Private SelectedNeoForge As DlNeoForgeListEntry = Nothing
    Private Sub SetNeoForgeInfoShow(IsShow As String)
        If PanNeoForgeInfo.Tag = IsShow Then Return
        PanNeoForgeInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanNeoForgeInfo, -CType(PanNeoForgeInfo.RenderTransform, TranslateTransform).Y, 200, 100, Ease:=New AniEaseOutFluent),
                AaOpacity(PanNeoForgeInfo, 1 - PanNeoForgeInfo.Opacity, 80, 90)
            }, "SetNeoForgeInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanNeoForgeInfo, 6 - CType(PanNeoForgeInfo.RenderTransform, TranslateTransform).Y, 120),
                AaOpacity(PanNeoForgeInfo, -PanNeoForgeInfo.Opacity, 50)
            }, "SetNeoForgeInfoShow")
        End If
    End Sub

    'Fabric
    Private SelectedFabric As String = Nothing
    Private Sub SetFabricInfoShow(IsShow As String)
        If PanFabricInfo.Tag = IsShow Then Return
        PanFabricInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanFabricInfo, -CType(PanFabricInfo.RenderTransform, TranslateTransform).Y, 200, 100, Ease:=New AniEaseOutFluent),
                AaOpacity(PanFabricInfo, 1 - PanFabricInfo.Opacity, 80, 90)
            }, "SetFabricInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanFabricInfo, 6 - CType(PanFabricInfo.RenderTransform, TranslateTransform).Y, 120),
                AaOpacity(PanFabricInfo, -PanFabricInfo.Opacity, 50)
            }, "SetFabricInfoShow")
        End If
    End Sub

    'FabricApi
    Private SelectedFabricApi As CompFile = Nothing
    Private Sub SetFabricApiInfoShow(IsShow As String)
        If PanFabricApiInfo.Tag = IsShow Then Return
        PanFabricApiInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanFabricApiInfo, -CType(PanFabricApiInfo.RenderTransform, TranslateTransform).Y, 200, 100, Ease:=New AniEaseOutFluent),
                AaOpacity(PanFabricApiInfo, 1 - PanFabricApiInfo.Opacity, 80, 90)
            }, "SetFabricApiInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanFabricApiInfo, 6 - CType(PanFabricApiInfo.RenderTransform, TranslateTransform).Y, 120),
                AaOpacity(PanFabricApiInfo, -PanFabricApiInfo.Opacity, 50)
            }, "SetFabricApiInfoShow")
        End If
    End Sub

    'Quilt
    Private SelectedQuilt As String = Nothing
    Private Sub SetQuiltInfoShow(IsShow As String)
        If PanQuiltInfo.Tag = IsShow Then Exit Sub
        PanQuiltInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanQuiltInfo, -CType(PanQuiltInfo.RenderTransform, TranslateTransform).Y, 270, 100, Ease:=New AniEaseOutBack),
                AaOpacity(PanQuiltInfo, 1 - PanQuiltInfo.Opacity, 100, 90)
            }, "SetQuiltInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanQuiltInfo, 6 - CType(PanQuiltInfo.RenderTransform, TranslateTransform).Y, 200),
                AaOpacity(PanQuiltInfo, -PanQuiltInfo.Opacity, 100)
            }, "SetQuiltInfoShow")
        End If
    End Sub

    'QSL
    Private SelectedQSL As CompFile = Nothing
    Private Sub SetQSLInfoShow(IsShow As String)
        If PanQSLInfo.Tag = IsShow Then Exit Sub
        PanQSLInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanQSLInfo, -CType(PanQSLInfo.RenderTransform, TranslateTransform).Y, 270, 100, Ease:=New AniEaseOutBack),
                AaOpacity(PanQSLInfo, 1 - PanQSLInfo.Opacity, 100, 90)
            }, "SetQSLInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanQSLInfo, 6 - CType(PanQSLInfo.RenderTransform, TranslateTransform).Y, 200),
                AaOpacity(PanQSLInfo, -PanQSLInfo.Opacity, 100)
            }, "SetQSLInfoShow")
        End If
    End Sub

    'LabyMod
    Private SelectedLabyModChannel As String = Nothing
    Private SelectedLabyModCommitRef As String = Nothing
    Private SelectedLabyModVersion As String = Nothing
    Private Sub SetLabyModInfoShow(IsShow As String)
        If PanLabyModInfo.Tag = IsShow Then Exit Sub
        PanLabyModInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanLabyModInfo, -CType(PanLabyModInfo.RenderTransform, TranslateTransform).Y, 270, 100, Ease:=New AniEaseOutBack),
                AaOpacity(PanLabyModInfo, 1 - PanLabyModInfo.Opacity, 100, 90)
            }, "SetLabyModInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanLabyModInfo, 6 - CType(PanLabyModInfo.RenderTransform, TranslateTransform).Y, 200),
                AaOpacity(PanLabyModInfo, -PanLabyModInfo.Opacity, 100)
            }, "SetLabyModInfoShow")
        End If
    End Sub

    'OptiFabric
    Private SelectedOptiFabric As CompFile = Nothing
    Private Sub SetOptiFabricInfoShow(IsShow As String)
        If PanOptiFabricInfo.Tag = IsShow Then Return
        PanOptiFabricInfo.Tag = IsShow
        If IsShow = "True" Then
            '显示信息栏
            AniStart({
                AaTranslateY(PanOptiFabricInfo, -CType(PanOptiFabricInfo.RenderTransform, TranslateTransform).Y, 200, 100, Ease:=New AniEaseOutFluent),
                AaOpacity(PanOptiFabricInfo, 1 - PanOptiFabricInfo.Opacity, 80, 90)
            }, "SetOptiFabricInfoShow")
        Else
            '隐藏信息栏
            AniStart({
                AaTranslateY(PanOptiFabricInfo, 6 - CType(PanOptiFabricInfo.RenderTransform, TranslateTransform).Y, 120),
                AaOpacity(PanOptiFabricInfo, -PanOptiFabricInfo.Opacity, 50)
            }, "SetOptiFabricInfoShow")
        End If
    End Sub

    Private IsReloading As Boolean = False '#3742 中，LoadOptiFineGetError 会初始化 LoadOptiFine，触发事件 LoadOptiFine.StateChanged，导致再次调用 SelectReload
    ''' <summary>
    ''' 重载已选择的项目的显示。
    ''' </summary>
    Private Sub SelectReload() Handles CardOptiFine.Swap, LoadOptiFine.StateChanged, CardForge.Swap, LoadForge.StateChanged, CardNeoForge.Swap, LoadNeoForge.StateChanged, CardFabric.Swap, LoadFabric.StateChanged, CardFabricApi.Swap, LoadFabricApi.StateChanged, CardOptiFabric.Swap, LoadOptiFabric.StateChanged, CardLiteLoader.Swap, LoadLiteLoader.StateChanged, LoadQuilt.StateChanged, CardQuilt.Swap, LoadQSL.StateChanged, CardQSL.Swap, LoadCleanroom.StateChanged, CardCleanroom.Swap, LoadLabyMod.StateChanged, CardLabyMod.Swap
        If SelectedMinecraftId Is Nothing OrElse IsReloading Then Return
        IsReloading = True
        '主预览
        SelectNameUpdate()
        ImgLogo.Source = GetSelectLogo()
        'OptiFine
        Dim OptiFineError As String = LoadOptiFineGetError()
        CardOptiFine.MainSwap.Visibility = If(OptiFineError Is Nothing, Visibility.Visible, Visibility.Collapsed)
        If OptiFineError IsNot Nothing Then CardOptiFine.IsSwaped = True '例如在同时展开卡片时选择了不兼容项则强制折叠
        SetOptiFineInfoShow(CardOptiFine.IsSwaped)
        If SelectedOptiFine Is Nothing Then
            BtnOptiFineClear.Visibility = Visibility.Collapsed
            ImgOptiFine.Visibility = Visibility.Collapsed
            LabOptiFine.Text = If(OptiFineError, GetLang("LangDownloadInstallClickToChose"))
            LabOptiFine.Foreground = ColorGray4
        Else
            BtnOptiFineClear.Visibility = Visibility.Visible
            ImgOptiFine.Visibility = Visibility.Visible
            LabOptiFine.Text = SelectedOptiFine.NameDisplay.Replace(SelectedMinecraftId & " ", "")
            LabOptiFine.Foreground = ColorGray1
        End If
        'LiteLoader
        If Not SelectedMinecraftId.Contains("1.") OrElse Val(SelectedMinecraftId.Split(".")(1)) > 12 Then
            CardLiteLoader.Visibility = Visibility.Collapsed
        Else
            CardLiteLoader.Visibility = Visibility.Visible
            Dim LiteLoaderError As String = LoadLiteLoaderGetError()
            CardLiteLoader.MainSwap.Visibility = If(LiteLoaderError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If LiteLoaderError IsNot Nothing Then CardLiteLoader.IsSwaped = True '例如在同时展开卡片时选择了不兼容项则强制折叠
            SetLiteLoaderInfoShow(CardLiteLoader.IsSwaped)
            If SelectedLiteLoader Is Nothing Then
                BtnLiteLoaderClear.Visibility = Visibility.Collapsed
                ImgLiteLoader.Visibility = Visibility.Collapsed
                LabLiteLoader.Text = If(LiteLoaderError, GetLang("LangDownloadInstallClickToChose"))
                LabLiteLoader.Foreground = ColorGray4
            Else
                BtnLiteLoaderClear.Visibility = Visibility.Visible
                ImgLiteLoader.Visibility = Visibility.Visible
                LabLiteLoader.Text = SelectedLiteLoader.Inherit
                LabLiteLoader.Foreground = ColorGray1
            End If
        End If
        'Forge
        Dim ForgeError As String = LoadForgeGetError()
        CardForge.MainSwap.Visibility = If(ForgeError Is Nothing, Visibility.Visible, Visibility.Collapsed)
        If ForgeError IsNot Nothing Then CardForge.IsSwaped = True
        SetForgeInfoShow(CardForge.IsSwaped)
        If SelectedForge Is Nothing Then
            BtnForgeClear.Visibility = Visibility.Collapsed
            ImgForge.Visibility = Visibility.Collapsed
            LabForge.Text = If(ForgeError, GetLang("LangDownloadInstallClickToChose"))
            LabForge.Foreground = ColorGray4
        Else
            BtnForgeClear.Visibility = Visibility.Visible
            ImgForge.Visibility = Visibility.Visible
            LabForge.Text = SelectedForge.VersionName
            LabForge.Foreground = ColorGray1
        End If
        'Cleanroom
        If SelectedMinecraftId = "1.12.2" Then
            CardCleanroom.Visibility = Visibility.Visible
            Dim CleanroomError As String = LoadCleanroomGetError()
            CardCleanroom.MainSwap.Visibility = If(CleanroomError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If CleanroomError IsNot Nothing Then CardCleanroom.IsSwaped = True
            SetCleanroomInfoShow(CardCleanroom.IsSwaped)
            If SelectedCleanroom Is Nothing Then
                BtnCleanroomClear.Visibility = Visibility.Collapsed
                ImgCleanroom.Visibility = Visibility.Collapsed
                LabCleanroom.Text = If(CleanroomError, "可以添加")
                LabCleanroom.Foreground = ColorGray4
            Else
                BtnCleanroomClear.Visibility = Visibility.Visible
                ImgCleanroom.Visibility = Visibility.Visible
                LabCleanroom.Text = SelectedCleanroom.VersionName
                LabCleanroom.Foreground = ColorGray1
            End If
        Else
            CardCleanroom.Visibility = Visibility.Collapsed
        End If
        'NeoForge
        If Not SelectedMinecraftId.Contains("1.") OrElse Val(SelectedMinecraftId.Split(".")(1)) <= 19 Then
            CardNeoForge.Visibility = Visibility.Collapsed
        Else
            CardNeoForge.Visibility = Visibility.Visible
            Dim NeoForgeError As String = LoadNeoForgeGetError()
            CardNeoForge.MainSwap.Visibility = If(NeoForgeError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If NeoForgeError IsNot Nothing Then CardNeoForge.IsSwaped = True
            SetNeoForgeInfoShow(CardNeoForge.IsSwaped)
            If SelectedNeoForge Is Nothing Then
                BtnNeoForgeClear.Visibility = Visibility.Collapsed
                ImgNeoForge.Visibility = Visibility.Collapsed
                LabNeoForge.Text = If(NeoForgeError, GetLang("LangDownloadInstallClickToChose"))
                LabNeoForge.Foreground = ColorGray4
            Else
                BtnNeoForgeClear.Visibility = Visibility.Visible
                ImgNeoForge.Visibility = Visibility.Visible
                LabNeoForge.Text = SelectedNeoForge.VersionName
                LabNeoForge.Foreground = ColorGray1
            End If
        End If
        'Fabric
        If SelectedMinecraftId.Contains("1.") AndAlso Val(SelectedMinecraftId.Split(".")(1)) <= 13 Then
            CardFabric.Visibility = Visibility.Collapsed
        Else
            CardFabric.Visibility = Visibility.Visible
            Dim FabricError As String = LoadFabricGetError()
            CardFabric.MainSwap.Visibility = If(FabricError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If FabricError IsNot Nothing Then CardFabric.IsSwaped = True
            SetFabricInfoShow(CardFabric.IsSwaped)
            If SelectedFabric Is Nothing Then
                BtnFabricClear.Visibility = Visibility.Collapsed
                ImgFabric.Visibility = Visibility.Collapsed
                LabFabric.Text = If(FabricError, GetLang("LangDownloadInstallClickToChose"))
                LabFabric.Foreground = ColorGray4
            Else
                BtnFabricClear.Visibility = Visibility.Visible
                ImgFabric.Visibility = Visibility.Visible
                LabFabric.Text = SelectedFabric.Replace("+build", "")
                LabFabric.Foreground = ColorGray1
            End If
        End If
        'FabricApi
        If SelectedFabric Is Nothing AndAlso SelectedQuilt Is Nothing Then
            CardFabricApi.Visibility = Visibility.Collapsed
        Else
            CardFabricApi.Visibility = Visibility.Visible
            Dim FabricApiError As String = LoadFabricApiGetError()
            CardFabricApi.MainSwap.Visibility = If(FabricApiError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If FabricApiError IsNot Nothing OrElse SelectedFabric Is Nothing AndAlso SelectedQuilt Is Nothing Then CardFabricApi.IsSwaped = True
            SetFabricApiInfoShow(CardFabricApi.IsSwaped)
            If SelectedFabricApi Is Nothing Then
                BtnFabricApiClear.Visibility = Visibility.Collapsed
                ImgFabricApi.Visibility = Visibility.Collapsed
                LabFabricApi.Text = If(FabricApiError, GetLang("LangDownloadInstallClickToChose"))
                LabFabricApi.Foreground = ColorGray4
            Else
                BtnFabricApiClear.Visibility = Visibility.Visible
                ImgFabricApi.Visibility = Visibility.Visible
                LabFabricApi.Text = SelectedFabricApi.DisplayName.Split("]")(1).Replace("Fabric API ", "").Replace(" build ", ".").Split("+").First.Trim
                LabFabricApi.Foreground = ColorGray1
            End If
        End If
        'Quilt
        If SelectedMinecraftId.Contains("1.") AndAlso Val(SelectedMinecraftId.Split(".")(1)) <= 14 AndAlso Not SelectedMinecraftId.Contains("1.14.4") Then
            CardQuilt.Visibility = Visibility.Collapsed
        Else
            CardQuilt.Visibility = Visibility.Visible
            Dim QuiltError As String = LoadQuiltGetError()
            CardQuilt.MainSwap.Visibility = If(QuiltError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If QuiltError IsNot Nothing Then CardQuilt.IsSwaped = True
            SetQuiltInfoShow(CardQuilt.IsSwaped)
            If SelectedQuilt Is Nothing Then
                BtnQuiltClear.Visibility = Visibility.Collapsed
                ImgQuilt.Visibility = Visibility.Collapsed
                LabQuilt.Text = If(QuiltError, "可以添加")
                LabQuilt.Foreground = ColorGray4
            Else
                BtnQuiltClear.Visibility = Visibility.Visible
                ImgQuilt.Visibility = Visibility.Visible
                LabQuilt.Text = SelectedQuilt.Replace("+build", "")
                LabQuilt.Foreground = ColorGray1
            End If
        End If
        'QSL
        If SelectedQuilt Is Nothing Then
            CardQSL.Visibility = Visibility.Collapsed
        Else
            CardQSL.Visibility = Visibility.Visible
            Dim QSLError As String = LoadQSLGetError()
            CardQSL.MainSwap.Visibility = If(QSLError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If QSLError IsNot Nothing OrElse SelectedQuilt Is Nothing Then CardQSL.IsSwaped = True
            SetQSLInfoShow(CardQSL.IsSwaped)
            If SelectedQSL Is Nothing Then
                BtnQSLClear.Visibility = Visibility.Collapsed
                ImgQSL.Visibility = Visibility.Collapsed
                LabQSL.Text = If(QSLError, "可以添加")
                LabQSL.Foreground = ColorGray4
            Else
                BtnQSLClear.Visibility = Visibility.Visible
                ImgQSL.Visibility = Visibility.Visible
                LabQSL.Text = SelectedQSL.DisplayName.Split("]")(1).Trim
                LabQSL.Foreground = ColorGray1
            End If
        End If
        'LabyMod
        If SelectedMinecraftId.Contains("1.") AndAlso Val(SelectedMinecraftId.Split(".")(1)) < 8 Then
            CardLabyMod.Visibility = Visibility.Collapsed
        Else
            CardLabyMod.Visibility = Visibility.Visible
            Dim LabyModError As String = LoadLabyModGetError()
            CardLabyMod.MainSwap.Visibility = If(LabyModError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If LabyModError IsNot Nothing Then CardLabyMod.IsSwaped = True
            SetLabyModInfoShow(CardLabyMod.IsSwaped)
            If SelectedLabyModVersion Is Nothing Then
                BtnLabyModClear.Visibility = Visibility.Collapsed
                ImgLabyMod.Visibility = Visibility.Collapsed
                LabLabyMod.Text = If(LabyModError, "可以添加")
                LabLabyMod.Foreground = ColorGray4
            Else
                BtnLabyModClear.Visibility = Visibility.Visible
                ImgLabyMod.Visibility = Visibility.Visible
                LabLabyMod.Text = SelectedLabyModVersion
                LabLabyMod.Foreground = ColorGray1
            End If
        End If
        'OptiFabric
        If SelectedFabric Is Nothing OrElse SelectedOptiFine Is Nothing Then
            CardOptiFabric.Visibility = Visibility.Collapsed
        Else
            CardOptiFabric.Visibility = Visibility.Visible
            Dim OptiFabricError As String = LoadOptiFabricGetError()
            CardOptiFabric.MainSwap.Visibility = If(OptiFabricError Is Nothing, Visibility.Visible, Visibility.Collapsed)
            If OptiFabricError IsNot Nothing OrElse SelectedFabric Is Nothing Then CardOptiFabric.IsSwaped = True
            SetOptiFabricInfoShow(CardOptiFabric.IsSwaped)
            If SelectedOptiFabric Is Nothing Then
                BtnOptiFabricClear.Visibility = Visibility.Collapsed
                ImgOptiFabric.Visibility = Visibility.Collapsed
                LabOptiFabric.Text = If(OptiFabricError, GetLang("LangDownloadInstallClickToChose"))
                LabOptiFabric.Foreground = ColorGray4
            Else
                BtnOptiFabricClear.Visibility = Visibility.Visible
                ImgOptiFabric.Visibility = Visibility.Visible
                LabOptiFabric.Text = SelectedOptiFabric.DisplayName.ToLower.Replace("optifabric-", "").Replace(".jar", "").Trim.TrimStart("v")
                LabOptiFabric.Foreground = ColorGray1
            End If
        End If
        '主警告
        If SelectedFabric IsNot Nothing AndAlso SelectedFabricApi Is Nothing Then
            HintFabricAPI.Visibility = Visibility.Visible
        Else
            HintFabricAPI.Visibility = Visibility.Collapsed
        End If
        If SelectedQuilt IsNot Nothing AndAlso SelectedQSL Is Nothing AndAlso SelectedFabricApi Is Nothing Then
            HintQSL.Visibility = Visibility.Visible
        Else
            HintQSL.Visibility = Visibility.Collapsed
        End If
        If SelectedQuilt IsNot Nothing AndAlso SelectedFabricApi IsNot Nothing AndAlso DlQSLLoader.Output IsNot Nothing Then
            For Each Version In DlQSLLoader.Output
                If IsSuitableQSL(Version.GameVersions, SelectedMinecraftId) Then
                    HintQuiltFabricAPI.Visibility = Visibility.Visible
                    Exit For
                Else
                    HintQuiltFabricAPI.Visibility = Visibility.Collapsed
                End If
            Next
        Else
            HintQuiltFabricAPI.Visibility = Visibility.Collapsed
        End If
        If SelectedFabric IsNot Nothing AndAlso SelectedOptiFine IsNot Nothing AndAlso SelectedOptiFabric Is Nothing Then
            If SelectedMinecraftId.StartsWith("1.14") OrElse SelectedMinecraftId.StartsWith("1.15") Then
                HintOptiFabric.Visibility = Visibility.Collapsed
                HintOptiFabricOld.Visibility = Visibility.Visible
            Else
                HintOptiFabric.Visibility = Visibility.Visible
                HintOptiFabricOld.Visibility = Visibility.Collapsed
            End If
        Else
            HintOptiFabric.Visibility = Visibility.Collapsed
            HintOptiFabricOld.Visibility = Visibility.Collapsed
        End If
        If SelectedMinecraftId.Contains("1.") AndAlso Val(SelectedMinecraftId.Split(".")(1)) >= 16 AndAlso SelectedOptiFine IsNot Nothing AndAlso
           (SelectedForge IsNot Nothing OrElse SelectedFabric IsNot Nothing) Then
            HintModOptiFine.Visibility = Visibility.Visible
        Else
            HintModOptiFine.Visibility = Visibility.Collapsed
        End If
        '结束
        IsReloading = False
    End Sub
    ''' <summary>
    ''' 清空已选择的项目。
    ''' </summary>
    Private Sub SelectClear()
        SelectedMinecraftId = Nothing
        SelectedMinecraftJsonUrl = Nothing
        SelectedMinecraftIcon = Nothing
        SelectedOptiFine = Nothing
        SelectedLiteLoader = Nothing
        SelectedLoaderName = Nothing
        SelectedAPIName = Nothing
        SelectedForge = Nothing
        SelectedNeoForge = Nothing
        SelectedCleanroom = Nothing
        SelectedFabric = Nothing
        SelectedFabricApi = Nothing
        SelectedQuilt = Nothing
        SelectedQSL = Nothing
        SelectedOptiFabric = Nothing
        SelectedLabyModCommitRef = Nothing
        SelectedLabyModVersion = Nothing
        SelectedLabyModChannel = Nothing
    End Sub

    ''' <summary>
    ''' 获取版本图标。
    ''' </summary>
    Private Function GetSelectLogo() As String
        If SelectedFabric IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/Fabric.png"
        ElseIf SelectedForge IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/Anvil.png"
        ElseIf SelectedNeoForge IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/NeoForge.png"
        ElseIf SelectedLiteLoader IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/Egg.png"
        ElseIf SelectedOptiFine IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/GrassPath.png"
        ElseIf SelectedQuilt IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/Quilt.png"
        ElseIf SelectedCleanroom IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/Cleanroom.png"
        ElseIf SelectedLabyModVersion IsNot Nothing Then
            Return "pack://application:,,,/images/Blocks/LabyMod.png"
        Else
            Return SelectedMinecraftIcon
        End If
    End Function

    '版本名处理
    ''' <summary>
    ''' 获取默认版本名。
    ''' </summary>
    Private Function GetSelectName() As String
        Dim Name As String = SelectedMinecraftId
        If SelectedFabric IsNot Nothing Then
            Name += "-Fabric_" & SelectedFabric.Replace("+build", "")
        End If
        If SelectedQuilt IsNot Nothing Then
            Name += "-Quilt_" & SelectedQuilt
        End If
        If SelectedLabyModVersion IsNot Nothing Then
            Name += "-LabyMod_" & SelectedLabyModVersion.Replace(" 稳定版", "_Production").Replace(" 快照版", "_Snapshot")
        End If
        If SelectedForge IsNot Nothing Then
            Name += "-Forge_" & SelectedForge.VersionName
        End If
        If SelectedNeoForge IsNot Nothing Then
            Name += "-NeoForge_" & SelectedNeoForge.VersionName
        End If
        If SelectedCleanroom IsNot Nothing Then
            Name += "-Cleanroom_" & SelectedCleanroom.VersionName
        End If
        If SelectedLiteLoader IsNot Nothing Then
            Name += "-LiteLoader"
        End If
        If SelectedOptiFine IsNot Nothing Then
            Name += "-OptiFine_" & SelectedOptiFine.NameDisplay.Replace(SelectedMinecraftId & " ", "").Replace(" ", "_")
        End If
        Return Name
    End Function
    Private IsSelectNameEdited As Boolean = False
    Private IsSelectNameChanging As Boolean = False
    Private Sub SelectNameUpdate()
        If IsSelectNameEdited OrElse IsSelectNameChanging Then Return
        IsSelectNameChanging = True
        TextSelectName.Text = GetSelectName()
        IsSelectNameChanging = False
    End Sub
    Private Sub TextSelectName_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TextSelectName.TextChanged
        If IsSelectNameChanging Then Return
        IsSelectNameEdited = True
        SelectReload()
    End Sub
    Private Sub TextSelectName_ValidateChanged(sender As Object, e As EventArgs) Handles TextSelectName.ValidateChanged
        BtnStart.IsEnabled = TextSelectName.IsValidated
    End Sub

#End Region

#Region "加载器"

    '结果数据化
    Private Sub LoadMinecraft_OnFinish()
        ExitSelectPage() '返回
        Try
            Dim Dict As New Dictionary(Of String, List(Of JObject)) From {
                {"正式版", New List(Of JObject)},
                {"预览版", New List(Of JObject)},
                {"远古版", New List(Of JObject)},
                {"愚人节版", New List(Of JObject)}
            }
            Dim Versions As JArray = DlClientListLoader.Output.Value("versions")
            For Each Version As JObject In Versions
                '确定分类
                Dim Type As String = Version("type").ToString()
                Dim versionId = Version("id").ToString().ToLower()
                Select Case Type
                    Case "release"
                        Type = "正式版"
                    Case "snapshot", "pending"
                        Type = "预览版"
                        'Mojang 误分类
                        If versionId.StartsWith("1.") AndAlso
                            Not versionId.Contains("combat") AndAlso
                            Not versionId.Contains("rc") AndAlso
                            Not versionId.Contains("experimental") AndAlso
                            Not versionId.Equals("1.2") AndAlso
                            Not versionId.Contains("pre") Then
                            Type = "正式版"
                            Version("type") = "release"
                        End If
                        '愚人节版本
                        Select Case Version("id").ToString.ToLower
                            Case "2point0_blue", "2point0_red", "2point0_purple", "2.0_blue", "2.0_red", "2.0_purple", "2.0"
                                Type = "愚人节版"
                                Version("id") = Version("id").ToString().Replace("point", ".")
                                Version("type") = "special"
                                Version.Add("lore", GetMcFoolName(Version("id")))
                            Case "20w14infinite", "20w14∞"
                                Type = "愚人节版"
                                Version("id") = "20w14∞"
                                Version("type") = "special"
                                Version.Add("lore", GetMcFoolName(Version("id")))
                            Case "3d shareware v1.34", "1.rv-pre1", "15w14a", "2.0", "22w13oneblockatatime", "23w13a_or_b", "24w14potato", "25w14craftmine"
                                Type = "愚人节版"
                                Version("type") = "special"
                                Version.Add("lore", GetMcFoolName(Version("id")))
                            Case Else '4/1 自动视作愚人节版
                                Dim ReleaseDate = Version("releaseTime").Value(Of Date).ToUniversalTime().AddHours(2)
                                If ReleaseDate.Month = 4 AndAlso ReleaseDate.Day = 1 Then
                                    Type = "愚人节版"
                                    Version("type") = "special"
                                End If
                        End Select
                    Case "special"
                        '已被处理的愚人节版
                        Type = "愚人节版"
                    Case Else
                        Type = "远古版"
                End Select
                '加入辞典
                Dict(Type).Add(Version)
            Next
            '排序
            For Each Pair In Dict.ToList
                Dict(Pair.Key) = Pair.Value.OrderByDescending(Function(j) j("releaseTime").Value(Of Date)).ToList
            Next
            '清空当前
            PanMinecraft.Children.Clear()
            '添加最新版本
            Dim CardInfo As New MyCard With {.Title = GetLang("LangDownloadLatest"), .Margin = New Thickness(0, 15, 0, 15)}
            Dim TopestVersions As New List(Of JObject)
            Dim Release As JObject = Dict("正式版")(0).DeepClone()
            Release("lore") = GetLang("LangDownloadClientReleaseReleaseOn") & " " & GetLocalTimeFormat(Release("releaseTime").Value(Of Date))
            TopestVersions.Add(Release)
            If Dict("正式版")(0)("releaseTime").Value(Of Date) < Dict("预览版")(0)("releaseTime").Value(Of Date) Then
                Dim Snapshot As JObject = Dict("预览版")(0).DeepClone()
                Snapshot("lore") = GetLang("LangDownloadClientBetaReleaseOn") & " " & GetLocalTimeFormat(Snapshot("releaseTime").Value(Of Date))
                TopestVersions.Add(Snapshot)
            End If
            Dim PanInfo As New StackPanel With {.Margin = New Thickness(20, MyCard.SwapedHeight, 18, 0), .VerticalAlignment = VerticalAlignment.Top, .RenderTransform = New TranslateTransform(0, 0), .Tag = TopestVersions}
            Dim StackInstall = Sub(Stack As StackPanel)
                               For Each item In Stack.Tag
                                   Stack.Children.Add(McDownloadListItem(item, Sub(sender, e) FrmDownloadInstall.MinecraftSelected(sender, e), False))
                               Next
                           End Sub
            MyCard.StackInstall(PanInfo, StackInstall)
            CardInfo.Children.Add(PanInfo)
            PanMinecraft.Children.Insert(0, CardInfo)
            '添加其他版本
            For Each Pair As KeyValuePair(Of String, List(Of JObject)) In Dict
                If Not Pair.Value.Any() Then Continue For
                '增加卡片
                Dim NewCard As New MyCard With {.Title = GetLangByWord(Pair.Key) & " (" & Pair.Value.Count & ")", .Margin = New Thickness(0, 0, 0, 15)}
                Dim NewStack As New StackPanel With {.Margin = New Thickness(20, MyCard.SwapedHeight, 18, 0), .VerticalAlignment = VerticalAlignment.Top, .RenderTransform = New TranslateTransform(0, 0), .Tag = Pair.Value}
                NewCard.Children.Add(NewStack)
                NewCard.SwapControl = NewStack
                '不能使用 AddressOf，这导致了 #535，原因完全不明，疑似是编译器 Bug
                NewCard.InstallMethod = StackInstall
                NewCard.IsSwaped = True
                PanMinecraft.Children.Add(NewCard)
            Next
            '自动选择版本
            If McVersionWaitingForSelect Is Nothing Then Exit Try
            Log("[Download] 自动选择 MC 版本：" & McVersionWaitingForSelect)
            For Each Version As JObject In Versions
                If Version("id").ToString <> McVersionWaitingForSelect Then Continue For
                Dim Item = McDownloadListItem(Version, Sub()
                                                       End Sub, False)
                MinecraftSelected(Item, Nothing)
            Next
        Catch ex As Exception
            Log(ex, "可视化安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub
    ''' <summary>
    ''' 当 MC 版本列表加载完时，立即自动选择的版本。用于外部调用。
    ''' </summary>
    Public Shared McVersionWaitingForSelect As String = Nothing

#End Region

#Region "OptiFine 列表"

    ''' <summary>
    ''' 获取 OptiFine 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadOptiFineGetError() As String
        If SelectedLoaderName = "NeoForge" OrElse SelectedLoaderName = "Quilt" OrElse SelectedLoaderName = "LabyMod" Then Return $"与 {SelectedLoaderName} 不兼容"
        If LoadOptiFine Is Nothing OrElse LoadOptiFine.State.LoadingState = MyLoading.MyLoadingState.Run Then Return "加载中……"
        If LoadOptiFine.State.LoadingState = MyLoading.MyLoadingState.Error Then Return "获取版本列表失败：" & CType(LoadOptiFine.State, Object).Error.Message
        '是否有 Cleanroom
        If SelectedCleanroom IsNot Nothing Then Return "与 Cleanroom 不兼容"
        '检查 Forge 1.13 - 1.14.3：全部不兼容
        If SelectedLoaderName = "Forge" AndAlso
            VersionSortInteger(SelectedMinecraftId, "1.13") >= 0 AndAlso VersionSortInteger("1.14.3", SelectedMinecraftId) >= 0 Then
            Return GetLang("LangDownloadInstallForgeIncompatible")
        End If
        '检查 Forge 版本
        Dim HasAny As Boolean = False
        Dim HasRequiredVersion As Boolean = False
        For Each OptiFineVersion As DlOptiFineListEntry In DlOptiFineListLoader.Output.Value
            If Not OptiFineVersion.NameDisplay.StartsWith(SelectedMinecraftId & " ") Then Continue For '不是同一个大版本
            HasAny = True
            If SelectedForge Is Nothing Then Return Nothing '未选择 Forge
            If IsOptiFineSuitForForge(OptiFineVersion, SelectedForge) Then Return Nothing '该版本可用
            If OptiFineVersion.RequiredForgeVersion IsNot Nothing Then HasRequiredVersion = True
        Next
        If Not HasAny Then
            Return GetLang("LangDownloadInstallNoAvailableVersion")
        ElseIf HasRequiredVersion Then
            Return GetLang("LangDownloadInstallForgeCompatibleTargetVersion")
        Else
            Return GetLang("LangDownloadInstallForgeIncompatible")
        End If
    End Function

    '检查某个 OptiFine 是否与某个 Forge 兼容
    Private Function IsOptiFineSuitForForge(OptiFine As DlOptiFineListEntry, Forge As DlForgeVersionEntry)
        If Forge.Inherit <> OptiFine.Inherit Then Return False '不是同一个大版本
        If OptiFine.RequiredForgeVersion Is Nothing Then Return False '不兼容 Forge
        If String.IsNullOrWhiteSpace(OptiFine.RequiredForgeVersion) Then Return True '#4183
        If OptiFine.RequiredForgeVersion.Contains(".") Then 'XX.X.XXX
            Return VersionSortInteger(Forge.Version.ToString, OptiFine.RequiredForgeVersion) = 0
        Else 'XXXX
            Return Forge.Version.Revision = OptiFine.RequiredForgeVersion
        End If
    End Function

    '限制展开
    Private Sub CardOptiFine_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardOptiFine.PreviewSwap
        If LoadOptiFineGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 OptiFine 版本列表。
    ''' </summary>
    Private Sub OptiFine_Loaded() Handles LoadOptiFine.StateChanged
        Try
            If DlOptiFineListLoader.State <> LoadState.Finished Then Return

            '获取版本列表
            Dim Versions As New List(Of DlOptiFineListEntry)
            For Each Version As DlOptiFineListEntry In DlOptiFineListLoader.Output.Value
                If SelectedForge IsNot Nothing AndAlso Not IsOptiFineSuitForForge(Version, SelectedForge) Then Continue For
                If Version.NameDisplay.StartsWith(SelectedMinecraftId & " ") Then Versions.Add(Version)
            Next
            If Not Versions.Any() Then Return
            '排序
            Versions = Versions.Sort(
            Function(Left As DlOptiFineListEntry, Right As DlOptiFineListEntry) As Boolean
                If Not Left.IsPreview AndAlso Right.IsPreview Then Return True
                If Left.IsPreview AndAlso Not Right.IsPreview Then Return False
                Return VersionSortBoolean(Left.NameDisplay, Right.NameDisplay)
            End Function)
            '可视化
            PanOptiFine.Children.Clear()
            For Each Version In Versions
                PanOptiFine.Children.Add(OptiFineDownloadListItem(Version, AddressOf OptiFine_Selected, False))
            Next
        Catch ex As Exception
            Log(ex, "可视化 OptiFine 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub OptiFine_Selected(sender As MyListItem, e As EventArgs)
        SelectedOptiFine = sender.Tag
        If SelectedForge IsNot Nothing AndAlso Not IsOptiFineSuitForForge(SelectedOptiFine, SelectedForge) Then SelectedForge = Nothing
        OptiFabric_Loaded()
        Forge_Loaded()
        NeoForge_Loaded()
        CardOptiFine.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub OptiFine_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnOptiFineClear.MouseLeftButtonUp
        SelectedOptiFine = Nothing
        SelectedOptiFabric = Nothing
        AutoSelectedOptiFabric = False
        CardOptiFine.IsSwaped = True
        e.Handled = True
        Forge_Loaded()
        NeoForge_Loaded()
        SelectReload()
    End Sub

#End Region

#Region "LiteLoader 列表"

    ''' <summary>
    ''' 获取 LiteLoader 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadLiteLoaderGetError() As String
        If Not SelectedMinecraftId.Contains("1.") OrElse Val(SelectedMinecraftId.Split(".")(1)) > 12 Then Return GetLang("LangDownloadInstallNoAvailableVersion")
        If LoadLiteLoader Is Nothing OrElse LoadLiteLoader.State.LoadingState = MyLoading.MyLoadingState.Run Then Return GetLang("LangDownloadInstallGettingList")
        If LoadLiteLoader.State.LoadingState = MyLoading.MyLoadingState.Error Then Return GetLang("LangDownloadInstallFailGetList") & CType(LoadLiteLoader.State, Object).Error.Message
        For Each Version As DlLiteLoaderListEntry In DlLiteLoaderListLoader.Output.Value
            If Version.Inherit = SelectedMinecraftId Then Return Nothing
        Next
        Return GetLang("LangDownloadInstallNoAvailableVersion")
    End Function

    '限制展开
    Private Sub CardLiteLoader_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardLiteLoader.PreviewSwap
        If LoadLiteLoaderGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 LiteLoader 版本列表。
    ''' </summary>
    Private Sub LiteLoader_Loaded() Handles LoadLiteLoader.StateChanged
        Try
            If DlLiteLoaderListLoader.State <> LoadState.Finished Then Return
            '获取版本列表
            Dim Versions As New List(Of DlLiteLoaderListEntry)
            For Each Version As DlLiteLoaderListEntry In DlLiteLoaderListLoader.Output.Value
                If Version.Inherit = SelectedMinecraftId Then Versions.Add(Version)
            Next
            If Not Versions.Any() Then Return
            '可视化
            PanLiteLoader.Children.Clear()
            For Each Version In Versions
                PanLiteLoader.Children.Add(LiteLoaderDownloadListItem(Version, AddressOf LiteLoader_Selected, False))
            Next
        Catch ex As Exception
            Log(ex, "可视化 LiteLoader 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub LiteLoader_Selected(sender As MyListItem, e As EventArgs)
        SelectedLiteLoader = sender.Tag
        CardLiteLoader.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub LiteLoader_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnLiteLoaderClear.MouseLeftButtonUp
        SelectedLiteLoader = Nothing
        CardLiteLoader.IsSwaped = True
        e.Handled = True
        SelectReload()
    End Sub

#End Region

#Region "Forge 列表"

    ''' <summary>
    ''' 获取 Forge 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadForgeGetError() As String
        If Not SelectedMinecraftId.StartsWith("1.") Then Return GetLang("LangDownloadInstallNoAvailableVersion")
        If Not LoadForge.State.IsLoader Then Return GetLang("LangDownloadInstallGettingList")
        Dim Loader As LoaderTask(Of String, List(Of DlForgeVersionEntry)) = LoadForge.State
        If SelectedMinecraftId <> Loader.Input Then Return GetLang("LangDownloadInstallGettingList")
        If Loader.State = LoadState.Loading Then Return GetLang("LangDownloadInstallGettingList")
        If Loader.State = LoadState.Failed Then
            Dim ErrorMessage As String = Loader.Error.Message
            If ErrorMessage.Contains(GetLang("LangDownloadInstallNoAvailableVersion")) Then
                Return GetLang("LangDownloadInstallNoAvailableVersion")
            Else
                Return GetLang("LangDownloadInstallFailGetList") & ErrorMessage
            End If
        End If
        If Loader.State <> LoadState.Finished Then Return GetLang("LangDownloadInstallFailGetListUnknownStatus") & " " & GetStringFromEnum(Loader.State)
        Dim NotSuitForOptiFine As Boolean = False
        For Each Version In Loader.Output
            If Version.Category = "universal" OrElse Version.Category = "client" Then Continue For '跳过无法自动安装的版本
            If SelectedLoaderName IsNot Nothing AndAlso SelectedLoaderName IsNot "Forge" Then Return $"与 {SelectedLoaderName} 不兼容"
            If SelectedOptiFine IsNot Nothing AndAlso
                VersionSortInteger(SelectedMinecraftId, "1.13") >= 0 AndAlso VersionSortInteger("1.14.3", SelectedMinecraftId) >= 0 Then
                Return GetLang("LangDownloadInstallOptiFineIncompatible") '1.13 ~ 1.14.3 OptiFine 检查
            End If
            If SelectedOptiFine IsNot Nothing AndAlso Not IsOptiFineSuitForForge(SelectedOptiFine, Version) Then
                NotSuitForOptiFine = True '与 OptiFine 不兼容
                Continue For
            End If
            Return Nothing
        Next
        Return If(NotSuitForOptiFine, GetLang("LangDownloadInstallOptiFineIncompatible"), GetLang("LangDownloadInstallNoAutoInstall"))
    End Function

    '限制展开
    Private Sub CardForge_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardForge.PreviewSwap
        If LoadForgeGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 Forge 版本列表。
    ''' </summary>
    Private Sub Forge_Loaded() Handles LoadForge.StateChanged
        Try
            If Not LoadForge.State.IsLoader Then Return
            Dim Loader As LoaderTask(Of String, List(Of DlForgeVersionEntry)) = LoadForge.State
            If SelectedMinecraftId <> Loader.Input Then Return
            If Loader.State <> LoadState.Finished Then Return
            '获取要显示的版本
            Dim Versions = Loader.Output.ToList '复制数组，以免 Output 在实例化后变空
            If Not Loader.Output.Any() Then Return
            PanForge.Children.Clear()
            Versions = Versions.Where(
            Function(v)
                If v.Category = "universal" OrElse v.Category = "client" Then Return False '跳过无法自动安装的版本
                If SelectedOptiFine IsNot Nothing AndAlso Not IsOptiFineSuitForForge(SelectedOptiFine, v) Then Return False
                Return True
            End Function).OrderByDescending(Function(v) v.Version).ToList()
            ForgeDownloadListItemPreload(PanForge, Versions, AddressOf Forge_Selected, False)
            For Each Version In Versions
                PanForge.Children.Add(ForgeDownloadListItem(Version, AddressOf Forge_Selected, False))
            Next
        Catch ex As Exception
            Log(ex, "可视化 Forge 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub Forge_Selected(sender As MyListItem, e As EventArgs)
        SelectedForge = sender.Tag
        SelectedLoaderName = "Forge"
        CardForge.IsSwaped = True
        If SelectedOptiFine IsNot Nothing AndAlso Not IsOptiFineSuitForForge(SelectedOptiFine, SelectedForge) Then SelectedOptiFine = Nothing
        OptiFine_Loaded()
        SelectReload()
    End Sub
    Private Sub Forge_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnForgeClear.MouseLeftButtonUp
        SelectedForge = Nothing
        SelectedLoaderName = Nothing
        CardForge.IsSwaped = True
        e.Handled = True
        OptiFine_Loaded()
        SelectReload()
    End Sub

#End Region

#Region "NeoForge 列表"

    ''' <summary>
    ''' 获取 NeoForge 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadNeoForgeGetError() As String
        If Not SelectedMinecraftId.StartsWith("1.") Then Return "不可用"
        If SelectedOptiFine IsNot Nothing Then Return "与 OptiFine 不兼容"
        If SelectedLoaderName IsNot Nothing AndAlso SelectedLoaderName IsNot "NeoForge" Then Return $"与 {SelectedLoaderName} 不兼容"
        If LoadNeoForge Is Nothing OrElse LoadNeoForge.State.LoadingState = MyLoading.MyLoadingState.Run Then Return "加载中……"
        If LoadNeoForge.State.LoadingState = MyLoading.MyLoadingState.Error Then Return "获取版本列表失败：" & CType(LoadNeoForge.State, Object).Error.Message
        If DlNeoForgeListLoader.Output.Value.Any(Function(v) v.Inherit = SelectedMinecraftId) Then
            Return Nothing
        Else
            Return GetLang("LangDownloadInstallNoAvailableVersion")
        End If
    End Function

    '限制展开
    Private Sub CardNeoForge_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardNeoForge.PreviewSwap
        If LoadNeoForgeGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 NeoForge 版本列表。
    ''' </summary>
    Private Sub NeoForge_Loaded() Handles LoadNeoForge.StateChanged
        Try
            '获取版本列表
            If DlNeoForgeListLoader.State <> LoadState.Finished Then Return
            Dim Versions = DlNeoForgeListLoader.Output.Value.Where(Function(v) v.Inherit = SelectedMinecraftId).ToList
            If Not Versions.Any() Then Return
            '可视化
            PanNeoForge.Children.Clear()
            NeoForgeDownloadListItemPreload(PanNeoForge, Versions, AddressOf NeoForge_Selected, False)
            For Each Version In Versions
                PanNeoForge.Children.Add(NeoForgeDownloadListItem(Version, AddressOf NeoForge_Selected, False))
            Next
        Catch ex As Exception
            Log(ex, "可视化 NeoForge 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub NeoForge_Selected(sender As MyListItem, e As EventArgs)
        SelectedNeoForge = sender.Tag
        SelectedLoaderName = "NeoForge"
        CardNeoForge.IsSwaped = True
        OptiFine_Loaded()
        SelectReload()
    End Sub
    Private Sub NeoForge_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnNeoForgeClear.MouseLeftButtonUp
        SelectedNeoForge = Nothing
        SelectedLoaderName = Nothing
        CardNeoForge.IsSwaped = True
        e.Handled = True
        OptiFine_Loaded()
        SelectReload()
    End Sub

#End Region

#Region "Cleanroom 列表"

    ''' <summary>
    ''' 获取 Cleanroom 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadCleanroomGetError() As String
        If Not SelectedMinecraftId.StartsWith("1.") Then Return "没有可用版本"
        If SelectedOptiFine IsNot Nothing Then Return "与 OptiFine 不兼容"
        If SelectedLoaderName IsNot Nothing AndAlso SelectedLoaderName IsNot "Cleanroom" Then Return $"与 {SelectedLoaderName} 不兼容"
        If LoadCleanroom Is Nothing OrElse LoadCleanroom.State.LoadingState = MyLoading.MyLoadingState.Run Then Return "正在获取版本列表……"
        If LoadCleanroom.State.LoadingState = MyLoading.MyLoadingState.Error Then Return "获取版本列表失败：" & CType(LoadCleanroom.State, Object).Error.Message
        Return Nothing
        'If DlCleanroomListLoader.Output.Value.Any(Function(v) v.Inherit = SelectedMinecraftId) Then
        '    Return Nothing
        'Else
        '    Return "没有可用版本"
        'End If
    End Function

    '限制展开
    Private Sub CardCleanroom_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardCleanroom.PreviewSwap
        If LoadCleanroomGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 Cleanroom 版本列表。
    ''' </summary>
    Private Sub Cleanroom_Loaded() Handles LoadCleanroom.StateChanged
        Try
            '获取版本列表
            If DlCleanroomListLoader.State <> LoadState.Finished Then Exit Sub
            Dim Versions = DlCleanroomListLoader.Output.Value.Where(Function(v) v.Inherit = SelectedMinecraftId).ToList
            If Not Versions.Any() Then Exit Sub
            '可视化
            PanCleanroom.Children.Clear()
            CleanroomDownloadListItemPreload(PanCleanroom, Versions, AddressOf Cleanroom_Selected, False)
            For Each Version In Versions
                PanCleanroom.Children.Add(CleanroomDownloadListItem(Version, AddressOf Cleanroom_Selected, False))
            Next
        Catch ex As Exception
            Log(ex, "可视化 Cleanroom 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub Cleanroom_Selected(sender As MyListItem, e As EventArgs)
        SelectedCleanroom = sender.Tag
        SelectedLoaderName = "Cleanroom"
        CardCleanroom.IsSwaped = True
        OptiFine_Loaded()
        SelectReload()
    End Sub
    Private Sub Cleanroom_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnCleanroomClear.MouseLeftButtonUp
        SelectedCleanroom = Nothing
        SelectedLoaderName = Nothing
        CardCleanroom.IsSwaped = True
        e.Handled = True
        OptiFine_Loaded()
        SelectReload()
    End Sub

#End Region

#Region "Fabric 列表"

    ''' <summary>
    ''' 获取 Fabric 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadFabricGetError() As String
        If LoadFabric Is Nothing OrElse LoadFabric.State.LoadingState = MyLoading.MyLoadingState.Run Then Return GetLang("LangDownloadInstallGettingList")
        If LoadFabric.State.LoadingState = MyLoading.MyLoadingState.Error Then Return GetLang("LangDownloadInstallFailGetList") & CType(LoadFabric.State, Object).Error.Message
        For Each Version As JObject In DlFabricListLoader.Output.Value("game")
            If Version("version").ToString = SelectedMinecraftId.Replace("∞", "infinite").Replace("Combat Test 7c", "1.16_combat-3") Then
                If SelectedLoaderName IsNot Nothing AndAlso SelectedLoaderName IsNot "Fabric" Then Return $"与 {SelectedLoaderName} 不兼容"
                Return Nothing
            End If
        Next
        Return GetLang("LangDownloadInstallNoAvailableVersion")
    End Function

    '限制展开
    Private Sub CardFabric_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardFabric.PreviewSwap
        If LoadFabricGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 Fabric 版本列表。
    ''' </summary>
    Private Sub Fabric_Loaded() Handles LoadFabric.StateChanged
        Try
            If DlFabricListLoader.State <> LoadState.Finished Then Return
            '获取版本列表
            Dim Versions As JArray = DlFabricListLoader.Output.Value("loader")
            If Not Versions.Any() Then Return
            '可视化
            PanFabric.Children.Clear()
            PanFabric.Tag = Versions
            CardFabric.SwapControl = PanFabric
            CardFabric.InstallMethod = Sub(Stack As StackPanel)
                                           For Each item In Stack.Tag
                                               Stack.Children.Add(FabricDownloadListItem(CType(item, JObject), AddressOf FrmDownloadInstall.Fabric_Selected))
                                           Next
                                       End Sub
        Catch ex As Exception
            Log(ex, "可视化 Fabric 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Public Sub Fabric_Selected(sender As MyListItem, e As EventArgs)
        SelectedFabric = sender.Tag("version").ToString
        SelectedLoaderName = "Fabric"
        FabricApi_Loaded()
        OptiFabric_Loaded()
        CardFabric.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub Fabric_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnFabricClear.MouseLeftButtonUp
        SelectedFabric = Nothing
        SelectedFabricApi = Nothing
        AutoSelectedFabricApi = False
        SelectedOptiFabric = Nothing
        AutoSelectedOptiFabric = False
        SelectedLoaderName = Nothing
        SelectedAPIName = Nothing
        CardFabric.IsSwaped = True
        e.Handled = True
        SelectReload()
    End Sub

#End Region

#Region "Fabric API 列表"

    ''' <summary>
    ''' 从显示名判断该 API 是否与某版本适配。
    ''' </summary>
    Public Shared Function IsSuitableFabricApi(DisplayName As String, MinecraftVersion As String) As Boolean
        Try
            If DisplayName Is Nothing OrElse MinecraftVersion Is Nothing Then Return False
            DisplayName = DisplayName.ToLower : MinecraftVersion = MinecraftVersion.Replace("∞", "infinite").Replace("Combat Test 7c", "1.16_combat-3").ToLower
            If DisplayName.StartsWith("[" & MinecraftVersion & "]") Then Return True
            If Not DisplayName.Contains("/") OrElse Not DisplayName.Contains("]") Then Return False
            '直接的判断（例如 1.18.1/22w03a）
            For Each Part As String In DisplayName.BeforeFirst("]").TrimStart("[").Split("/")
                If Part = MinecraftVersion Then Return True
            Next
            '将版本名分割语素（例如 1.16.4/5）
            Dim Lefts = RegexSearch(DisplayName.BeforeFirst("]"), "[a-z/]+|[0-9/]+")
            Dim Rights = RegexSearch(MinecraftVersion.BeforeFirst("]"), "[a-z/]+|[0-9/]+")
            '对每段进行判断
            Dim i As Integer = 0
            While True
                '两边均缺失，感觉是一个东西
                If Lefts.Count - 1 < i AndAlso Rights.Count - 1 < i Then Return True
                '确定两边是否一致
                Dim LeftValue As String = If(Lefts.Count - 1 < i, "-1", Lefts(i))
                Dim RightValue As String = If(Rights.Count - 1 < i, "-1", Rights(i))
                If Not LeftValue.Contains("/") Then
                    If LeftValue <> RightValue Then Return False
                Else
                    '左边存在斜杠
                    If Not LeftValue.Contains(RightValue) Then Return False
                End If
                i += 1
            End While
            Return True
        Catch ex As Exception
            Log(ex, "判断 Fabric API 版本适配性出错（" & DisplayName & ", " & MinecraftVersion & "）")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取 FabricApi 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadFabricApiGetError() As String
        If LoadFabricApi Is Nothing OrElse LoadFabricApi.State.LoadingState = MyLoading.MyLoadingState.Run Then Return GetLang("LangDownloadInstallGettingList")
        If LoadFabricApi.State.LoadingState = MyLoading.MyLoadingState.Error Then Return GetLang("LangDownloadInstallFailGetList") & CType(LoadFabricApi.State, Object).Error.Message
        If SelectedAPIName IsNot Nothing AndAlso SelectedAPIName IsNot "Fabric API" Then Return $"与 {SelectedAPIName} 不兼容"
        If DlFabricApiLoader.Output Is Nothing Then
            If SelectedFabric Is Nothing AndAlso SelectedQuilt Is Nothing Then Return "需要安装 Fabric / Quilt"
            Return "加载中……"
        End If
        For Each Version In DlFabricApiLoader.Output
            If Not IsSuitableFabricApi(Version.DisplayName, SelectedMinecraftId) Then Continue For
            If SelectedFabric Is Nothing AndAlso SelectedQuilt Is Nothing Then Return "需要安装 Fabric / Quilt"
            Return Nothing
        Next
        Return GetLang("LangDownloadInstallNoAvailableVersion")
    End Function

    '限制展开
    Private Sub CardFabricApi_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardFabricApi.PreviewSwap
        If LoadFabricApiGetError() IsNot Nothing Then e.Handled = True
    End Sub

    Private AutoSelectedFabricApi As Boolean = False
    ''' <summary>
    ''' 尝试重新可视化 FabricApi 版本列表。
    ''' </summary>
    Private Sub FabricApi_Loaded() Handles LoadFabricApi.StateChanged
        Try
            If DlFabricApiLoader.State <> LoadState.Finished Then Exit Sub
            If SelectedMinecraftId Is Nothing OrElse (SelectedFabric Is Nothing AndAlso SelectedQuilt Is Nothing) Then Exit Sub
            '获取版本列表
            Dim Versions As New List(Of CompFile)
            For Each Version In DlFabricApiLoader.Output
                If IsSuitableFabricApi(Version.DisplayName, SelectedMinecraftId) Then
                    If Not Version.DisplayName.StartsWith("[") Then
                        Log("[Download] 已特判修改 Fabric API 显示名：" & Version.DisplayName, LogLevel.Debug)
                        Version.DisplayName = "[" & SelectedMinecraftId & "] " & Version.DisplayName
                    End If
                    Versions.Add(Version)
                End If
            Next
            If Not Versions.Any() Then Return
            Versions = Versions.OrderByDescending(Function(v) v.ReleaseDate).ToList
            '可视化
            PanFabricApi.Children.Clear()
            For Each Version In Versions
                If Not IsSuitableFabricApi(Version.DisplayName, SelectedMinecraftId) Then Continue For
                PanFabricApi.Children.Add(FabricApiDownloadListItem(Version, AddressOf FabricApi_Selected))
            Next
            '自动选择 Fabric API
            If (Not AutoSelectedFabricApi AndAlso SelectedQuilt Is Nothing) OrElse (SelectedQuilt IsNot Nothing AndAlso LoadQSLGetError() Is "没有可用版本") Then
                AutoSelectedFabricApi = True
                Log($"[Download] 已自动选择 Fabric API：{CType(PanFabricApi.Children(0), MyListItem).Title}")
                FabricApi_Selected(PanFabricApi.Children(0), Nothing)
            End If
        Catch ex As Exception
            Log(ex, "可视化 Fabric API 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub FabricApi_Selected(sender As MyListItem, e As EventArgs)
        SelectedFabricApi = sender.Tag
        SelectedAPIName = "Fabric API"
        CardFabricApi.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub FabricApi_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnFabricApiClear.MouseLeftButtonUp
        SelectedFabricApi = Nothing
        SelectedAPIName = Nothing
        CardFabricApi.IsSwaped = True
        e.Handled = True
        SelectReload()
    End Sub

#End Region

#Region "Quilt 列表"

    ''' <summary>
    ''' 获取 Quilt 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadQuiltGetError() As String
        If LoadQuilt Is Nothing OrElse LoadQuilt.State.LoadingState = MyLoading.MyLoadingState.Run Then Return "加载中……"
        If LoadQuilt.State.LoadingState = MyLoading.MyLoadingState.Error Then Return "获取版本列表失败：" & CType(LoadQuilt.State, Object).Error.Message
        For Each Version As JObject In DlQuiltListLoader.Output.Value("game")
            If Version("version").ToString = SelectedMinecraftId.Replace("∞", "infinite").Replace("Combat Test 7c", "1.16_combat-3") Then
                If SelectedOptiFine IsNot Nothing Then Return "与 OptiFine 不兼容"
                If SelectedLoaderName IsNot Nothing AndAlso SelectedLoaderName IsNot "Quilt" Then Return $"与 {SelectedLoaderName} 不兼容"
                Return Nothing
            End If
        Next
        Return "不可用"
    End Function

    '限制展开
    Private Sub CardQuilt_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardQuilt.PreviewSwap
        If LoadQuiltGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 Quilt 版本列表。
    ''' </summary>
    Private Sub Quilt_Loaded() Handles LoadQuilt.StateChanged
        Try
            If DlQuiltListLoader.State <> LoadState.Finished Then Exit Sub
            '获取版本列表
            Dim Versions As JArray = DlQuiltListLoader.Output.Value("loader")
            If Not Versions.Any() Then Exit Sub
            '可视化
            PanQuilt.Children.Clear()
            PanQuilt.Tag = Versions
            CardQuilt.SwapControl = PanQuilt
            CardQuilt.InstallMethod = Sub(Stack As StackPanel)
                                          For Each item In Stack.Tag
                                              Stack.Children.Add(QuiltDownloadListItem(CType(item, JObject), AddressOf FrmDownloadInstall.Quilt_Selected))
                                          Next
                                      End Sub
        Catch ex As Exception
            Log(ex, "可视化 Quilt 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Public Sub Quilt_Selected(sender As MyListItem, e As EventArgs)
        SelectedQuilt = sender.Tag("version").ToString
        SelectedLoaderName = "Quilt"
        FabricApi_Loaded()
        QSL_Loaded()
        CardQuilt.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub Quilt_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnQuiltClear.MouseLeftButtonUp
        SelectedQuilt = Nothing
        SelectedQSL = Nothing
        SelectedFabricApi = Nothing
        SelectedLoaderName = Nothing
        SelectedAPIName = Nothing
        CardQuilt.IsSwaped = True
        e.Handled = True
        SelectReload()
    End Sub

#End Region

#Region "QSL 列表"

    ''' <summary>
    ''' 从显示名判断该 API 是否与某版本适配。
    ''' </summary>
    Public Shared Function IsSuitableQSL(SupportVersions As List(Of String), MinecraftVersion As String) As Boolean
        Try
            If SupportVersions.Contains(MinecraftVersion) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Log(ex, "判断 QSL 版本适配性出错（" & SupportVersions.ToString & ", " & MinecraftVersion & "）")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取 QSL 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadQSLGetError() As String
        If LoadQSL Is Nothing OrElse LoadQSL.State.LoadingState = MyLoading.MyLoadingState.Run Then Return "正在获取版本列表……"
        If LoadQSL.State.LoadingState = MyLoading.MyLoadingState.Error Then Return "获取版本列表失败：" & CType(LoadQSL.State, Object).Error.Message
        If SelectedAPIName IsNot Nothing AndAlso SelectedAPIName IsNot "QFAPI / QSL" Then Return $"与 {SelectedAPIName} 不兼容"
        If DlQSLLoader.Output Is Nothing Then
            If SelectedQuilt Is Nothing Then Return "需要安装 Quilt"
            Return "正在获取版本列表……"
        End If
        For Each Version In DlQSLLoader.Output
            If Not IsSuitableQSL(Version.GameVersions, SelectedMinecraftId) Then Continue For
            If SelectedQuilt Is Nothing Then Return "需要安装 Quilt"
            Return Nothing
        Next
        Return "没有可用版本"
    End Function

    '限制展开
    Private Sub CardQSL_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardQSL.PreviewSwap
        If LoadQSLGetError() IsNot Nothing Then e.Handled = True
    End Sub

    Private AutoSelectedQSL As Boolean = False
    ''' <summary>
    ''' 尝试重新可视化 QSL 版本列表。
    ''' </summary>
    Private Sub QSL_Loaded() Handles LoadQSL.StateChanged
        Try
            If DlQSLLoader.State <> LoadState.Finished Then Exit Sub
            If SelectedMinecraftId Is Nothing OrElse SelectedQuilt Is Nothing Then Exit Sub
            '获取版本列表
            Dim Versions As New List(Of CompFile)
            For Each Version In DlQSLLoader.Output
                If IsSuitableQSL(Version.GameVersions, SelectedMinecraftId) Then
                    If Not Version.DisplayName.StartsWith("[") Then
                        Log("[Download] 已特判修改 QSL 显示名：" & Version.DisplayName, LogLevel.Debug)
                        Version.DisplayName = "[" & SelectedMinecraftId & "] " & Version.DisplayName
                    End If
                    Versions.Add(Version)
                End If
            Next
            If Not Versions.Any() Then Exit Sub
            Versions = Sort(Versions, Function(a, b) a.ReleaseDate > b.ReleaseDate)
            '可视化
            PanQSL.Children.Clear()
            For Each Version In Versions
                If Not IsSuitableQSL(Version.GameVersions, SelectedMinecraftId) Then Continue For
                PanQSL.Children.Add(QSLDownloadListItem(Version, AddressOf QSL_Selected))
            Next
            '自动选择 QSL
            If Not AutoSelectedQSL Then
                AutoSelectedQSL = True
                Log($"[Download] 已自动选择 QSL：{CType(PanQSL.Children(0), MyListItem).Title}")
                QSL_Selected(PanQSL.Children(0), Nothing)
            End If
        Catch ex As Exception
            Log(ex, "可视化 QSL 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub QSL_Selected(sender As MyListItem, e As EventArgs)
        SelectedQSL = sender.Tag
        SelectedAPIName = "QFAPI / QSL"
        CardQSL.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub QSL_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnQSLClear.MouseLeftButtonUp
        SelectedQSL = Nothing
        SelectedAPIName = Nothing
        CardQSL.IsSwaped = True
        e.Handled = True
        SelectReload()
    End Sub

#End Region

#Region "OptiFabric 列表"

    ''' <summary>
    ''' 从显示名判断该 Mod 是否与某版本适配。
    ''' </summary>
    Private Function IsSuitableOptiFabric(ModFile As CompFile, MinecraftVersion As String) As Boolean
        Try
            If MinecraftVersion Is Nothing Then Return False
            Return ModFile.GameVersions.Contains(MinecraftVersion)
        Catch ex As Exception
            Log(ex, "判断 OptiFabric 版本适配性出错（" & MinecraftVersion & "）")
            Return False
        End Try
    End Function

    Private AutoSelectedOptiFabric As Boolean = False
    ''' <summary>
    ''' 获取 OptiFabric 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadOptiFabricGetError() As String
        If SelectedMinecraftId.StartsWith("1.14") OrElse SelectedMinecraftId.StartsWith("1.15") Then Return GetLang("LangDownloadInstallOptiFabricManually")
        If LoadOptiFabric Is Nothing OrElse LoadOptiFabric.State.LoadingState = MyLoading.MyLoadingState.Run Then Return GetLang("LangDownloadInstallGettingList")
        If LoadOptiFabric.State.LoadingState = MyLoading.MyLoadingState.Error Then Return GetLang("LangDownloadInstallFailGetList") & CType(LoadOptiFabric.State, Object).Error.Message
        If DlOptiFabricLoader.Output Is Nothing Then
            If SelectedFabric Is Nothing AndAlso SelectedOptiFine Is Nothing Then Return GetLang("LangDownloadInstallOptiFineFabricNeed")
            If SelectedFabric Is Nothing Then Return GetLang("LangDownloadInstallFabricNeed")
            If SelectedOptiFine Is Nothing Then Return GetLang("LangDownloadInstallOptiFineNeed")
            Return GetLang("LangDownloadInstallGettingList")
        End If
        For Each Version In DlOptiFabricLoader.Output
            If Not IsSuitableOptiFabric(Version, SelectedMinecraftId) Then Continue For '2135#
            If SelectedFabric Is Nothing AndAlso SelectedOptiFine Is Nothing Then Return GetLang("LangDownloadInstallOptiFineFabricNeed")
            If SelectedFabric Is Nothing Then Return GetLang("LangDownloadInstallFabricNeed")
            If SelectedOptiFine Is Nothing Then Return GetLang("LangDownloadInstallOptiFineNeed")
            Return Nothing '通过检查
        Next
        Return GetLang("LangDownloadInstallNoAvailableVersion")
    End Function

    '限制展开
    Private Sub CardOptiFabric_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardOptiFabric.PreviewSwap
        If LoadOptiFabricGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 OptiFabric 版本列表。
    ''' </summary>
    Private Sub OptiFabric_Loaded() Handles LoadOptiFabric.StateChanged
        Try
            If DlOptiFabricLoader.State <> LoadState.Finished Then Return
            If SelectedMinecraftId Is Nothing OrElse SelectedFabric Is Nothing OrElse SelectedOptiFine Is Nothing Then Return
            '获取版本列表
            Dim Versions As New List(Of CompFile)
            For Each Version In DlOptiFabricLoader.Output
                If IsSuitableOptiFabric(Version, SelectedMinecraftId) Then Versions.Add(Version)
            Next
            If Not Versions.Any() Then Return
            '排序
            Versions = Versions.OrderByDescending(Function(v) v.ReleaseDate).ToList
            '可视化
            PanOptiFabric.Children.Clear()
            For Each Version In Versions
                If Not IsSuitableOptiFabric(Version, SelectedMinecraftId) Then Continue For
                PanOptiFabric.Children.Add(OptiFabricDownloadListItem(Version, AddressOf OptiFabric_Selected))
            Next
            '自动选择 OptiFabric
            If Not AutoSelectedOptiFabric AndAlso
                Not (SelectedMinecraftId.StartsWith("1.14") OrElse SelectedMinecraftId.StartsWith("1.15")) Then '1.14~15 不自动选择
                AutoSelectedOptiFabric = True
                Log($"[Download] 已自动选择 OptiFabric：{CType(PanOptiFabric.Children(0), MyListItem).Title}")
                OptiFabric_Selected(PanOptiFabric.Children(0), Nothing)
            End If
        Catch ex As Exception
            Log(ex, "可视化 OptiFabric 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Private Sub OptiFabric_Selected(sender As MyListItem, e As EventArgs)
        SelectedOptiFabric = sender.Tag
        CardOptiFabric.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub OptiFabric_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnOptiFabricClear.MouseLeftButtonUp
        SelectedOptiFabric = Nothing
        CardOptiFabric.IsSwaped = True
        e.Handled = True
        SelectReload()
    End Sub

#End Region

#Region "LabyMod 列表"

    ''' <summary>
    ''' 获取 LabyMod 的加载异常信息。若正常则返回 Nothing。
    ''' </summary>
    Private Function LoadLabyModGetError() As String
        If LoadLabyMod Is Nothing OrElse LoadLabyMod.State.LoadingState = MyLoading.MyLoadingState.Run Then Return "加载中……"
        If LoadLabyMod.State.LoadingState = MyLoading.MyLoadingState.Error Then Return "获取版本列表失败：" & CType(LoadLabyMod.State, Object).Error.Message
        For Each Version As JObject In DlLabyModListLoader.Output.Value("production")("minecraftVersions")
            If Version("version").ToString = SelectedMinecraftId Then
                If SelectedOptiFine IsNot Nothing Then Return "与 OptiFine 不兼容"
                If SelectedLoaderName IsNot Nothing AndAlso SelectedLoaderName IsNot "LabyMod" Then Return $"与 {SelectedLoaderName} 不兼容"
                Return Nothing
            End If
        Next
        For Each Version As JObject In DlLabyModListLoader.Output.Value("snapshot")("minecraftVersions")
            If Version("version").ToString = SelectedMinecraftId Then
                If SelectedOptiFine IsNot Nothing Then Return "与 OptiFine 不兼容"
                If SelectedLoaderName IsNot Nothing AndAlso SelectedLoaderName IsNot "LabyMod" Then Return $"与 {SelectedLoaderName} 不兼容"
                Return Nothing
            End If
        Next
        Return "不可用"
    End Function

    '限制展开
    Private Sub CardLabyMod_PreviewSwap(sender As Object, e As RouteEventArgs) Handles CardLabyMod.PreviewSwap
        If LoadLabyModGetError() IsNot Nothing Then e.Handled = True
    End Sub

    ''' <summary>
    ''' 尝试重新可视化 LabyMod 版本列表。
    ''' </summary>
    Private Sub LabyMod_Loaded() Handles LoadLabyMod.StateChanged
        Try
            If LoadLabyMod.State.LoadingState = MyLoading.MyLoadingState.Run Then Exit Sub
            '获取版本列表
            Dim Versions As JObject = DlLabyModListLoader.Output.Value
            If Versions Is Nothing OrElse Versions("production") Is Nothing OrElse Versions("snapshot") Is Nothing Then Exit Sub
            '可视化
            Dim ProcessedVersions As New JArray
            For Each Production As JObject In Versions("production")("minecraftVersions")
                If Production("version").ToString = SelectedMinecraftId Then
                    Dim ProductionVersion As New JObject
                    ProductionVersion.Add("version", Versions("production")("labyModVersion"))
                    ProductionVersion.Add("channel", "production")
                    ProductionVersion.Add("commitReference", Versions("production")("commitReference"))
                    ProcessedVersions.Add(ProductionVersion)
                End If
            Next
            For Each Snapshot As JObject In Versions("snapshot")("minecraftVersions")
                If Snapshot("version").ToString = SelectedMinecraftId Then
                    Dim SnapshotVersion As New JObject
                    SnapshotVersion.Add("version", Versions("production")("labyModVersion"))
                    SnapshotVersion.Add("channel", "snapshot")
                    SnapshotVersion.Add("commitReference", Versions("snapshot")("commitReference"))
                    ProcessedVersions.Add(SnapshotVersion)
                End If
            Next
            'MyMsgBox(If(ProcessedVersions.ToString, "Nothing"))
            PanLabyMod.Children.Clear()
            PanLabyMod.Tag = ProcessedVersions
            CardLabyMod.SwapControl = PanLabyMod
            CardLabyMod.InstallMethod = Sub(Stack As StackPanel)
                                            For Each item As JObject In Stack.Tag
                                                Stack.Children.Add(LabyModDownloadListItem(item, AddressOf FrmDownloadInstall.LabyMod_Selected))
                                            Next
                                        End Sub
        Catch ex As Exception
            Log(ex, "可视化 LabyMod 安装版本列表出错", LogLevel.Feedback)
        End Try
    End Sub

    '选择与清除
    Public Sub LabyMod_Selected(sender As MyListItem, e As EventArgs)
        SelectedLabyModChannel = sender.Tag("channel").ToString
        SelectedLabyModCommitRef = sender.Tag("commitReference").ToString
        SelectedLabyModVersion = sender.Tag("version").ToString & If(SelectedLabyModChannel = "snapshot", " 快照版", " 稳定版")
        SelectedLoaderName = "LabyMod"
        CardLabyMod.IsSwaped = True
        SelectReload()
    End Sub
    Private Sub LabyMod_Clear(sender As Object, e As MouseButtonEventArgs) Handles BtnLabyModClear.MouseLeftButtonUp
        SelectedLabyModCommitRef = Nothing
        SelectedLabyModVersion = Nothing
        SelectedLabyModChannel = Nothing
        SelectedLoaderName = Nothing
        SelectedAPIName = Nothing
        CardLabyMod.IsSwaped = True
        e.Handled = True
        SelectReload()
    End Sub
#End Region

#Region "安装"

    Private Sub TextSelectName_KeyDown(sender As Object, e As KeyEventArgs) Handles TextSelectName.KeyDown
        If e.Key = Key.Enter AndAlso BtnStart.IsEnabled Then BtnStart_Click()
    End Sub
    Private Sub BtnStart_Click() Handles BtnStart.Click
        '确认版本隔离
        If SelectedLoaderName IsNot Nothing AndAlso
           (Setup.Get("LaunchArgumentIndieV2") = 0 OrElse Setup.Get("LaunchArgumentIndieV2") = 2) Then
            If MyMsgBox("你尚未开启版本隔离，多个 MC 版本会共用同一个 Mod 文件夹。" & vbCrLf &
                        "因此，游戏可能会因为读取到与当前版本不符的 Mod 而崩溃。" & vbCrLf &
                        "推荐先在 设置 → 启动选项 → 默认版本隔离 中开启版本隔离！", "版本隔离提示", "取消下载", "继续") = 1 Then
                Return
            End If
        End If
        '提交安装申请
        Dim VersionName As String = TextSelectName.Text
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftJson = SelectedMinecraftJsonUrl,
            .MinecraftName = SelectedMinecraftId,
            .OptiFineEntry = SelectedOptiFine,
            .ForgeEntry = SelectedForge,
            .NeoForgeEntry = SelectedNeoForge,
            .CleanroomEntry = SelectedCleanroom,
            .FabricVersion = SelectedFabric,
            .FabricApi = SelectedFabricApi,
            .QuiltVersion = SelectedQuilt,
            .QSL = SelectedQSL,
            .OptiFabric = SelectedOptiFabric,
            .LiteLoaderEntry = SelectedLiteLoader,
            .LabyModChannel = SelectedLabyModChannel,
            .LabyModCommitRef = SelectedLabyModCommitRef
        }
        If Not McInstall(Request) Then Return
        '返回，这样在再次进入安装页面时这个版本就会显示文件夹已重复
        ExitSelectPage()
    End Sub

#End Region

End Class
