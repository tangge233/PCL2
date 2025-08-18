﻿Imports System.Windows.Threading
Imports System.IO.Compression
Imports PCL.Core.LifecycleManagement

Public Class Application

#If DEBUGRESERVED Then
    ''' <summary>
    ''' 用于开始程序时的一些测试。
    ''' </summary>
    Private Sub Test()
        Try
            ModDevelop.Start()
        Catch ex As Exception
            Log(ex, "开发者模式测试出错", LogLevel.Msgbox)
        End Try
    End Sub
#End If

    Public Sub New()
        '注册生命周期事件
        Lifecycle.When(LifecycleState.Loaded, AddressOf Application_Startup)
    End Sub

    '开始
    Private Sub Application_Startup() '(sender As Object, e As StartupEventArgs) Handles Me.Startup
        '刷新语言
        Try
            Application.Current.Resources.MergedDictionaries(1) = New ResourceDictionary With {.Source = New Uri("pack://application:,,,/Resources/Language/" & Lang & ".xaml", UriKind.RelativeOrAbsolute)}
        Catch ex As Exception
            MsgBox("无法找到语言资源：" & Lang & vbCrLf & "Language resource cannot be found:" & Lang, MsgBoxStyle.Critical)
            Lang = GetDefaultLang()
            WriteReg("Lang", Lang)
        End Try

        '依照选择语言切换字体
        Dim LaunchFont As FontFamily
        Select Case Lang
            Case "zh-TW", "zh-HK", "lzh", "zh-MARS"
                LaunchFont = New FontFamily(New Uri("pack://application:,,,/"), "./Resources/#PCL English, Segoe UI, Microsoft JhengHei UI")
            Case "ja-JP"
                LaunchFont = New FontFamily(New Uri("pack://application:,,,/"), "./Resources/#PCL English, Segoe UI, Yu Gothic UI, Microsoft YaHei UI")
            Case "ko-KR"
                LaunchFont = New FontFamily(New Uri("pack://application:,,,/"), "./Resources/#PCL English, Segoe UI, Malgun Gothic, Microsoft YaHei UI")
            Case "en-US", "en-GB", "zh-CN", "zh-MEME"
                LaunchFont = New FontFamily(New Uri("pack://application:,,,/"), "./Resources/#PCL English, Segoe UI, Microsoft YaHei UI")
            Case Else '非英语的其他西欧语言统一使用 Segoe UI
                LaunchFont = New FontFamily(New Uri("pack://application:,,,/"), "Segoe UI, ./Resources/#PCL English, Microsoft YaHei UI")
        End Select
        SwitchApplicationFont(LaunchFont)

        Try
            '创建自定义跟踪监听器，用于检测是否存在 Binding 失败
            PresentationTraceSources.DataBindingSource.Listeners.Add(New BindingErrorTraceListener())
            PresentationTraceSources.DataBindingSource.Switch.Level = SourceLevels.Error
            SecretOnApplicationStart()
            '检查参数调用
            Dim args = Environment.GetCommandLineArgs.Skip(1).ToArray()
            If args.Length > 0 Then
                If args(0) = "--update" Then
                    '自动更新
                    UpdateReplace(args(1), args(2).Trim(""""), args(3).Trim(""""), args(4))
                    Environment.Exit(ProcessReturnValues.TaskDone)
                ElseIf args(0) = "--gpu" Then
                    '调整显卡设置
                    Try
                        SetGPUPreference(args(1).Trim(""""))
                        Environment.Exit(ProcessReturnValues.TaskDone)
                    Catch ex As Exception
                        Environment.Exit(ProcessReturnValues.Fail)
                    End Try
                ElseIf args(0).StartsWithF("--memory") Then
                    '内存优化
                    Dim Ram = My.Computer.Info.AvailablePhysicalMemory
                    Try
                        PageOtherTest.MemoryOptimizeInternal(False)
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, GetLang("LangApplicationDialogContentMemReduceFail"))
                        Environment.Exit(-1)
                    End Try
                    If My.Computer.Info.AvailablePhysicalMemory < Ram Then '避免 ULong 相减出现负数
                        Environment.Exit(0)
                    Else
                        Environment.Exit((My.Computer.Info.AvailablePhysicalMemory - Ram) / 1024) '返回清理的内存量（K）
                    End If
#If DEBUGRESERVED Then
                    '制作更新包
                ElseIf args(0) = "--edit1" Then
                    ExeEdit(args(1), True)
                    Environment.Exit(ProcessReturnValues.TaskDone)
                ElseIf args(0) = "--edit2" Then
                    ExeEdit(args(1), False)
                    Environment.Exit(ProcessReturnValues.TaskDone)
#End If
                End If
            End If
            '初始化文件结构
            Directory.CreateDirectory(Path & "PCL\Pictures")
            Directory.CreateDirectory(Path & "PCL\Musics")
            Try
                Directory.CreateDirectory(PathTemp)
                If Not CheckPermission(PathTemp) Then Throw New Exception(GetLang("LangApplicationExceptionNoAccessPermission", PathTemp))
            Catch ex As Exception
                If PathTemp = IO.Path.GetTempPath() & "PCL\" Then
                    MyMsgBox(GetLang("LangApplicationDialogContentCacheFolderUnavailable", GetExceptionDetail(ex)), GetLang("LangApplicationDialogTitleCacheFolderUnavailable"))
                Else
                    MyMsgBox(GetLang("LangApplicationDialogContentCustomCacheFolderUnavailable", GetExceptionDetail(ex)), GetLang("LangApplicationDialogTitleCacheFolderUnavailable"))
                    Setup.Set("SystemSystemCache", "")
                    PathTemp = IO.Path.GetTempPath() & "PCL\"
                End If
            End Try
            Directory.CreateDirectory(PathTemp & "Cache")
            Directory.CreateDirectory(PathTemp & "Download")
            Directory.CreateDirectory(PathAppdata)
            '检测单例
#If Not DEBUGRESERVED Then
            Dim ShouldWaitForExit As Boolean = args.Length > 0 AndAlso args(0) = "--wait" '要求等待已有的 PCL 退出
            Dim WaitRetryCount As Integer = 0
WaitRetry:
            Dim WindowHwnd As IntPtr = FindWindow(Nothing, "Plain Craft Launcher Community Edition ")
            If WindowHwnd = IntPtr.Zero Then FindWindow(Nothing, "Plain Craft Launcher 2 Community Edition ")
            If WindowHwnd <> IntPtr.Zero Then
                If ShouldWaitForExit AndAlso WaitRetryCount < 20 Then '至多等待 10 秒
                    WaitRetryCount += 1
                    Thread.Sleep(500)
                    GoTo WaitRetry
                End If
                '将已有的 PCL 窗口拖出来
                ShowWindowToTop(WindowHwnd)
                '播放提示音并退出
                Beep()
                Environment.[Exit](ProcessReturnValues.Cancel)
            End If
#End If
            '设置 ToolTipService 默认值
            ToolTipService.InitialShowDelayProperty.OverrideMetadata(GetType(DependencyObject), New FrameworkPropertyMetadata(300))
            ToolTipService.BetweenShowDelayProperty.OverrideMetadata(GetType(DependencyObject), New FrameworkPropertyMetadata(400))
            ToolTipService.ShowDurationProperty.OverrideMetadata(GetType(DependencyObject), New FrameworkPropertyMetadata(9999999))
            ToolTipService.PlacementProperty.OverrideMetadata(GetType(DependencyObject), New FrameworkPropertyMetadata(Primitives.PlacementMode.Bottom))
            ToolTipService.HorizontalOffsetProperty.OverrideMetadata(GetType(DependencyObject), New FrameworkPropertyMetadata(8.0))
            ToolTipService.VerticalOffsetProperty.OverrideMetadata(GetType(DependencyObject), New FrameworkPropertyMetadata(4.0))
            '设置初始窗口
            If Setup.Get("UiLauncherLogo") Then
                FrmStart = New SplashScreen("Images\icon.ico")
                FrmStart.Show(False, True)
            End If
            '添加日志
            Log($"[Start] 程序版本：{VersionBaseName} (Channel: {VersionBranchName},Code: {VersionCode}{If(CommitHash = "", "", $"，#{CommitHash}")})")
            Log($"[Start] 识别码：{UniqueAddress}")
            Log($"[Start] 程序路径：{PathWithName}")
            Log($"[Start] 系统版本：{Environment.OSVersion.Version}, 架构：{Runtime.InteropServices.RuntimeInformation.OSArchitecture}")
            Log($"[Start] 系统编码：{Encoding.Default.HeaderName} ({Encoding.Default.CodePage}, GBK={IsGBKEncoding})")
            Log($"[Start] 管理员权限：{IsAdmin()}")
            Log("[Location] 启动器语言：" & Lang)
            Log("[Location] 当前系统环境是否为中国大陆：" & IsLocationZH())
            '检测异常环境
            If Path.Contains(IO.Path.GetTempPath()) OrElse Path.Contains("AppData\Local\Temp\") Then
                MyMsgBox(GetLang("LangApplicationDialogContentRunInTemp"), GetLang("LangApplicationDialogTitleRunInTemp"), GetLang("LangDialogThemeUnlockGameAccept"), IsWarn:=True)
            End If
            If Is32BitSystem Then
                MyMsgBox(GetLang("LangApplicationDialogContent32BitWarn"), GetLang("LangApplicationDialogTitleRunInTemp"), GetLang("LangDialogThemeUnlockGameAccept"), IsWarn:=True)
            End If
            Dim IS_WINDOWS_MEET_REQUIRE As Boolean = Environment.OSVersion.Version.Major >= 10
            Dim IS_FRAMEWORK_MEET_REQUIRE As Boolean
            Using key = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32)
                Using ndpKey = key.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\")
                    If ndpKey IsNot Nothing AndAlso ndpKey.GetValue("Release") IsNot Nothing Then
                        Dim rt = ndpKey.GetValue("Release")
                        IS_FRAMEWORK_MEET_REQUIRE = Val(rt) >= 528040
                        Log($"[Runtime] 检测到运行时版本为 {Val(rt)}")
                    Else
                        Log("[Runtime] 检测不到运行时")
                    End If
                End Using
            End Using
            Dim ProblemList As New List(Of String)
            If Not IS_WINDOWS_MEET_REQUIRE Then ProblemList.Add("Windows 版本不满足最低要求，最低需要 Windows 10 20H2")
            If Not IS_FRAMEWORK_MEET_REQUIRE Then ProblemList.Add(".NET Framework 版本不满足要求，需要 .NET Framework 4.8.1")
            If ProblemList.Count <> 0 Then
                MyMsgBox("PCL CE 在启动时检测到环境问题：" & vbCrLf & vbCrLf &
                         ProblemList.Join(vbCrLf) & vbCrLf & vbCrLf &
                         "需要解决这些问题才能正常使用启动器……",
                        Button2:=If(IS_WINDOWS_MEET_REQUIRE, String.Empty, "升级系统"),
                        Button2Action:=Sub() OpenWebsite("https://www.microsoft.com/zh-cn/software-download/windows10"),
                        Button3:=If(IS_FRAMEWORK_MEET_REQUIRE, String.Empty, "安装框架"),
                        Button3Action:=Sub() OpenWebsite("https://dotnet.microsoft.com/zh-cn/download/dotnet-framework/thank-you/net481-offline-installer"))
            End If
            Dim IS_WINDOWS_MEET_REQUIRE As Boolean = Environment.OSVersion.Version.Major >= 10
            Dim IS_FRAMEWORK_MEET_REQUIRE As Boolean
            Using key = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32)
                Using ndpKey = key.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\")
                    If ndpKey IsNot Nothing AndAlso ndpKey.GetValue("Release") IsNot Nothing Then
                        Dim rt = ndpKey.GetValue("Release")
                        IS_FRAMEWORK_MEET_REQUIRE = Val(rt) >= 528040
                        Log($"[Runtime] 检测到运行时版本为 {Val(rt)}")
                    Else
                        Log("[Runtime] 检测不到运行时")
                    End If
                End Using
            End Using
            Dim ProblemList As New List(Of String)
            If Not IS_WINDOWS_MEET_REQUIRE Then ProblemList.Add("Windows 版本不满足最低要求，最低需要 Windows 10 20H2")
            If Not IS_FRAMEWORK_MEET_REQUIRE Then ProblemList.Add(".NET Framework 版本不满足要求，需要 .NET Framework 4.8.1")
            If ProblemList.Count <> 0 Then
                MyMsgBox("PCL CE 在启动时检测到环境问题：" & vbCrLf & vbCrLf &
                         ProblemList.Join(vbCrLf) & vbCrLf & vbCrLf &
                         "需要解决这些问题才能正常使用启动器……",
                        Button2:=If(IS_WINDOWS_MEET_REQUIRE, String.Empty, "升级系统"),
                        Button2Action:=Sub() OpenWebsite("https://www.microsoft.com/zh-cn/software-download/windows10"),
                        Button3:=If(IS_FRAMEWORK_MEET_REQUIRE, String.Empty, "安装框架"),
                        Button3Action:=Sub() OpenWebsite("https://dotnet.microsoft.com/zh-cn/download/dotnet-framework/thank-you/net481-offline-installer"))
            End If
            '设置初始化
            Setup.Load("SystemDebugMode")
            Setup.Load("SystemDebugAnim")
            Setup.Load("ToolDownloadThread")
            Setup.Load("ToolDownloadCert")
            Setup.Load("ToolDownloadSpeed")
            '删除旧日志
            For i = 1 To 5
                Dim oldLogFile = $"{Path}PCL\Log-CE{i}.log"
                If File.Exists(oldLogFile) Then File.Delete(oldLogFile)
            Next
            '释放资源
            Directory.CreateDirectory(PathPure & "CE")
            SetDllDirectory(PathPure & "CE")
            Dim WebpPath = $"{PathPure}CE\libwebp.dll"
            If Not File.Exists(WebpPath) Then WriteFile(WebpPath, GetResources("libwebp64"))
            WriteFile(PathPure & "CE\" & "msalruntime.zip", GetResources("msalruntime"))
            If Not File.Exists(PathPure & "CE\msalruntime.dll") Then
                If Directory.Exists(PathPure & "CE\runtimes") Then DeleteDirectory(PathPure & "CE\runtimes")
                Using fs = New FileStream(PathPure & "CE\" & "msalruntime.zip", FileMode.Open, FileAccess.Read)
                    Using fszip = New ZipArchive(fs, ZipArchiveMode.Read)
                        fszip.ExtractToDirectory(PathPure & "CE\")
                    End Using
                End Using
            End If
            'Pipe RPC 初始化
            StartEchoPipe()
            '设置字体
            Dim TargetFont As String = Setup.Get("UiFont")
            If Not String.IsNullOrEmpty(TargetFont) Then
                Try
                    Dim Font = Fonts.SystemFontFamilies.FirstOrDefault(Function(x) x.FamilyNames.Values.Contains(TargetFont))
                    If Font Is Nothing Then
                        Setup.Reset("UiFont")
                    Else
                        SetLaunchFont(TargetFont)
                    End If
                Catch ex As Exception
                    Log(ex, "字体加载失败", LogLevel.Hint)
                    Setup.Reset("UiFont")
                End Try
            End If
            '计时
            Log("[Start] 第一阶段加载用时：" & GetTimeTick() - ApplicationStartTick & " ms")
            ApplicationStartTick = GetTimeTick()
            '执行测试
#If DEBUGRESERVED Then
            Test()
#End If
            AniControlEnabled += 1
        Catch ex As Exception
            Dim FilePath As String = Nothing
            Try
                FilePath = PathWithName
            Catch
            End Try
            MsgBox(GetExceptionDetail(ex, True) & vbCrLf & "PCL 所在路径：" & If(String.IsNullOrEmpty(FilePath), "获取失败", FilePath), MsgBoxStyle.Critical, GetLang("LangApplicationDialogTitleInitError"))
            FormMain.EndProgramForce(ProcessReturnValues.Exception)
        End Try
    End Sub

    '结束
    Private Sub Application_SessionEnding(sender As Object, e As SessionEndingCancelEventArgs) Handles Me.SessionEnding
        FrmMain.EndProgram(False)
    End Sub

    '异常
    Private Sub Application_DispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs) Handles Me.DispatcherUnhandledException
        On Error Resume Next
        e.Handled = True
        If IsProgramEnded Then Return
        FeedbackInfo()
        Dim Detail As String = GetExceptionDetail(e.Exception, True)
        If Detail.Contains("System.Windows.Threading.Dispatcher.Invoke") OrElse Detail.Contains("MS.Internal.AppModel.ITaskbarList.HrInit") OrElse Detail.Contains("未能加载文件或程序集") OrElse
           Detail.Contains(".NET Framework") Then ' “自动错误判断” 的结果分析
            OpenWebsite("https://dotnet.microsoft.com/zh-cn/download/dotnet-framework/thank-you/net481-offline-installer")
            Log(e.Exception, GetLang("LangApplicationDialogContentNETWarn"), LogLevel.Critical, GetLang("LangApplicationDialogTitleNETWarn"))
        Else
            Log(e.Exception, GetLang("LangApplicationDialogContentUnknownError"), LogLevel.Critical, GetLang("LangApplicationDialogTitleUnknownError"))
        End If
    End Sub

    Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" (lpPathName As String) As Boolean


    '切换窗口

    '控件模板事件
    Private Sub MyIconButton_Click(sender As Object, e As EventArgs)
    End Sub

    Public Shared ShowingTooltips As New List(Of Border)
    Private Sub TooltipLoaded(sender As Border, e As EventArgs)
        ShowingTooltips.Add(sender)
    End Sub
    Private Sub TooltipUnloaded(sender As Border, e As RoutedEventArgs)
        ShowingTooltips.Remove(sender)
    End Sub

    ' 自定义监听器类
    Public Class BindingErrorTraceListener
        Inherits TraceListener

        Public Overrides Sub Write(message As String)
            Log($"警告，检测到 Binding 失败：{message}")
        End Sub

        Public Overrides Sub WriteLine(message As String)
            Log($"警告，检测到 Binding 失败：{message}")
        End Sub
    End Class

End Class
