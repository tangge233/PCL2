﻿Imports System.IO.Compression

Public Module ModModpack

    '触发整合包安装的外部接口
    ''' <summary>
    ''' 弹窗要求选择一个整合包文件并进行安装。
    ''' </summary>
    Public Sub ModpackInstall()
        Dim File As String = SelectFile(GetLang("LangModModpackSelectModpack") & "(*.rar;*.zip;*.mrpack)|*.rar;*.zip;*.mrpack", GetLang("LangModModpackSelectTitleModpack")) '选择整合包文件
        If String.IsNullOrEmpty(File) Then Return
        RunInThread(
        Sub()
            Try
                ModpackInstall(File)
            Catch ex As CancelledException
            Catch ex As Exception
                Log(ex, GetLang("LangModModpackExceptionManualInstallFail"), LogLevel.Msgbox)
            End Try
        End Sub)
    End Sub
    ''' <summary>
    ''' 构建并启动安装给定的整合包文件的加载器，并返回该加载器。若失败则抛出异常。
    ''' 必须在工作线程执行。
    ''' </summary>
    ''' <exception cref="CancelledException" />
    Public Function ModpackInstall(File As String, Optional VersionName As String = Nothing, Optional Logo As String = Nothing) As LoaderCombo(Of String)
        Log("[ModPack] 整合包安装请求：" & If(File, "null"))
        Dim Archive As ZipArchive = Nothing
        Dim ArchiveBaseFolder As String = ""
        Try
            '字符校验
            Dim TargetFolder As String = $"{PathMcFolder}versions\{VersionName}\"
            If TargetFolder.Contains("!") OrElse TargetFolder.Contains(";") Then Hint("游戏路径中不能含有感叹号或分号：" & TargetFolder, HintType.Critical) : Throw New CancelledException
            '获取整合包种类与关键 Json
            Dim PackType As Integer = -1
            Try
                Archive = New ZipArchive(New FileStream(File, FileMode.Open, FileAccess.Read, FileShare.Read))
                '从根目录判断整合包类型
                If Archive.GetEntry("mcbbs.packmeta") IsNot Nothing Then PackType = 3 : Exit Try 'MCBBS 整合包（优先于 manifest.json 判断）
                If Archive.GetEntry("mmc-pack.json") IsNot Nothing Then PackType = 2 : Exit Try 'MMC 整合包（优先于 manifest.json 判断，#4194）
                If Archive.GetEntry("modrinth.index.json") IsNot Nothing Then PackType = 4 : Exit Try 'Modrinth 整合包
                If Archive.GetEntry("manifest.json") IsNot Nothing Then
                    Dim Json As JObject = GetJson(ReadFile(Archive.GetEntry("manifest.json").Open, Encoding.UTF8))
                    If Json("addons") Is Nothing Then
                        PackType = 0 : Exit Try 'CurseForge 整合包
                    Else
                        PackType = 3 : Exit Try 'MCBBS 整合包
                    End If
                End If
                If Archive.GetEntry("modpack.json") IsNot Nothing Then PackType = 1 : Exit Try 'HMCL 整合包
                If Archive.GetEntry("modpack.zip") IsNot Nothing OrElse Archive.GetEntry("modpack.mrpack") IsNot Nothing Then PackType = 9 : Exit Try '带启动器的压缩包
                '从一级目录判断整合包类型
                For Each Entry In Archive.Entries
                    Dim FullNames As String() = Entry.FullName.Split("/")
                    ArchiveBaseFolder = FullNames(0) & "/"
                    '确定为一级目录下
                    If FullNames.Count <> 2 Then Continue For
                    '判断是否为关键文件
                    If FullNames(1) = "mcbbs.packmeta" Then PackType = 3 : Exit Try 'MCBBS 整合包（优先于 manifest.json 判断）
                    If FullNames(1) = "mmc-pack.json" Then PackType = 2 : Exit Try 'MMC 整合包（优先于 manifest.json 判断，#4194）
                    If FullNames(1) = "modrinth.index.json" Then PackType = 4 : Exit Try 'Modrinth 整合包
                    If FullNames(1) = "manifest.json" Then
                        Dim Json As JObject = GetJson(ReadFile(Entry.Open, Encoding.UTF8))
                        If Json("addons") Is Nothing Then
                            PackType = 0 : Exit Try 'CurseForge 整合包
                        Else
                            PackType = 3 : ArchiveBaseFolder = "overrides/" : Exit Try 'MCBBS 整合包
                        End If
                    End If
                    If FullNames(1) = "modpack.json" Then PackType = 1 : Exit Try 'HMCL 整合包
                    If FullNames(1) = "modpack.zip" OrElse FullNames(1) = "modpack.mrpack" Then PackType = 9 : Exit Try '带启动器的压缩包
                Next
            Catch ex As Exception
                If GetExceptionDetail(ex, True).Contains("Error.WinIOError") Then
                    Throw New Exception(GetLang("LangModModpackExceptionOpenFileFail"), ex)
                ElseIf File.EndsWithF(".rar", True) Then
                    Throw New Exception(GetLang("LangModModpackExceptionRarFormat"), ex)
                Else
                    Throw New Exception(GetLang("LangModModpackExceptionReadFileFail"), ex)
                End If
            End Try
            '执行对应的安装方法
            Select Case PackType
                Case 0
                    Log("[ModPack] 整合包种类：CurseForge")
                    Return InstallPackCurseForge(File, Archive, ArchiveBaseFolder, VersionName, Logo)
                Case 1
                    Log("[ModPack] 整合包种类：HMCL")
                    Return InstallPackHMCL(File, Archive, ArchiveBaseFolder)
                Case 2
                    Log("[ModPack] 整合包种类：MMC")
                    Return InstallPackMMC(File, Archive, ArchiveBaseFolder)
                Case 3
                    Log("[ModPack] 整合包种类：MCBBS")
                    Return InstallPackMCBBS(File, Archive, ArchiveBaseFolder, VersionName)
                Case 4
                    Log("[ModPack] 整合包种类：Modrinth")
                    Return InstallPackModrinth(File, Archive, ArchiveBaseFolder, VersionName, Logo)
                Case 9
                    Log("[ModPack] 整合包种类：带启动器的压缩包")
                    Return InstallPackLauncherPack(File, Archive, ArchiveBaseFolder)
                Case Else
                    Log("[ModPack] 整合包种类：未能识别，假定为压缩包")
                    Return InstallPackCompress(File, Archive)
            End Select
        Finally
            If Archive IsNot Nothing Then Archive.Dispose()
        End Try
    End Function

    Private Sub ExtractModpackFiles(InstallTemp As String, FileAddress As String, Loader As LoaderBase, ProgressIncrement As Double)
        '解压文件
        Dim RetryCount As Integer = 1
        Dim Encode = Encoding.GetEncoding("GB18030")
        Dim InitialProgress = Loader.Progress
        Try
Retry:
            Loader.Progress = InitialProgress
            DeleteDirectory(InstallTemp)
            ExtractFile(FileAddress, InstallTemp, Encode, ProgressIncrementHandler:=Sub(Delta) Loader.Progress += Delta * ProgressIncrement)
        Catch ex As Exception
            Log(ex, "第 " & RetryCount & " 次解压尝试失败")
            If TypeOf ex Is ArgumentException Then
                Encode = Encoding.UTF8
                Log("[ModPack] 已切换压缩包解压编码为 UTF8")
            End If
            '完全不知道为啥会出现文件正在被另一进程使用的问题，总之加个重试
            If RetryCount < 5 Then
                Thread.Sleep(RetryCount * 2000)
                If Loader IsNot Nothing AndAlso Loader.LoadingState <> MyLoading.MyLoadingState.Run Then Return
                RetryCount += 1
                GoTo Retry
            Else
                Throw New Exception("解压整合包文件失败", ex)
            End If
        End Try
        Loader.Progress = InitialProgress + ProgressIncrement
    End Sub
    ''' <summary>
    ''' 从整合包的 override 目录复制文件，同时设置 PCL 的配置文件与版本隔离。
    ''' 对路径末尾是否为 \ 没有要求。
    ''' </summary>
    Private Sub CopyOverrideDirectory(OverridesFolder As String, VersionFolder As String, Loader As LoaderBase, ProgressIncrement As Double)
        If Not OverridesFolder.EndsWithF("\") Then OverridesFolder += "\"
        If Not VersionFolder.EndsWithF("\") Then VersionFolder += "\"
        '复制文件
        If Directory.Exists(OverridesFolder) Then
            Log($"[ModPack] 处理整合包覆写文件夹：{OverridesFolder} → {VersionFolder}")
            CopyDirectory(OverridesFolder, VersionFolder, Sub(Delta) Loader.Progress += Delta * ProgressIncrement)
        Else
            Log($"[ModPack] 整合包中没有覆写文件夹：{OverridesFolder}")
            Loader.Progress += ProgressIncrement
        End If
        '设置 ini
        Dim OverridesIni As String = $"{OverridesFolder}PCL\Setup.ini"
        Dim VersionIni As String = $"{VersionFolder}PCL\Setup.ini"
        If File.Exists(OverridesIni) Then
            WriteIni(OverridesIni, "VersionArgumentIndie", 1) '开启版本隔离
            WriteIni(OverridesIni, "VersionArgumentIndieV2", True)
            CopyFile(OverridesIni, VersionIni) '覆写已有的 ini
        Else
            WriteIni(VersionIni, "VersionArgumentIndie", 1) '开启版本隔离
            WriteIni(VersionIni, "VersionArgumentIndieV2", True)
        End If
        IniClearCache(VersionIni) '重置缓存，避免被安装过程中写入的 ini 覆盖
    End Sub

#Region "CurseForge"
    Private Function InstallPackCurseForge(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String,
                                           Optional VersionName As String = Nothing, Optional Logo As String = Nothing) As LoaderCombo(Of String)

        '读取 Json 文件
        Dim Json As JObject
        Try
            Json = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "manifest.json").Open))
        Catch ex As Exception
            Throw New Exception(GetLang("LangModModpackExceptionCurseforgeModpackError"), ex)
        End Try
        If Json("minecraft") Is Nothing OrElse Json("minecraft")("version") Is Nothing Then Throw New Exception(GetLang("LangModModpackCurseforgeNoGameInfo"))

        '获取版本名
        If VersionName Is Nothing Then
            VersionName = If(Json("name"), "")
            Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
            If Validate.Validate(VersionName) <> "" Then VersionName = ""
            If VersionName = "" Then VersionName = MyMsgBoxInput(GetLang("LangModModpackInputInstanceName"), "", "", New ObjectModel.Collection(Of Validate) From {Validate})
            If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        End If

        '获取 Mod API 版本信息
        Dim ForgeVersion As String = Nothing
        Dim NeoForgeVersion As String = Nothing
        Dim FabricVersion As String = Nothing
        Dim QuiltVersion As String = Nothing
        For Each Entry In If(Json("minecraft")("modLoaders"), {})
            Dim Id As String = If(Entry("id"), "").ToString.ToLower
            If Id.StartsWithF("forge-") Then
                'Forge 指定
                If Id.Contains("recommended") Then Throw New Exception(GetLang("LangModModpackExceptionModpackTooOld"))
                Log("[ModPack] 整合包 Forge 版本：" & Id)
                ForgeVersion = Id.Replace("forge-", "")
            ElseIf Id.StartsWithF("neoforge-") Then
                'NeoForge 指定
                Log("[ModPack] 整合包 NeoForge 版本：" & Id)
                NeoForgeVersion = Id.Replace("neoforge-", "")
            ElseIf Id.StartsWithF("fabric-") Then
                'Fabric 指定
                Try
                    Log("[ModPack] 整合包 Fabric 版本：" & Id)
                    FabricVersion = Id.Replace("fabric-", "")
                    Exit For
                Catch ex As Exception
                    Log(ex, "读取整合包 Fabric 版本失败：" & Id)
                End Try
            ElseIf Id.StartsWithF("quilt-") Then
                'Quilt 指定
                Try
                    Log("[ModPack] 整合包 Quilt 版本：" & Id)
                    QuiltVersion = Id.Replace("quilt-", "")
                    Exit For
                Catch ex As Exception
                    Log(ex, "读取整合包 Quilt 版本失败：" & Id)
                End Try
            End If
        Next
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim InstallLoaders As New List(Of LoaderBase)
        Dim OverrideHome As String = If(Json("overrides"), "")
        If OverrideHome <> "" Then
            InstallLoaders.Add(New LoaderTask(Of String, Integer)(GetLang("LangModModpackTaskExtraModpack"),
            Sub(Task As LoaderTask(Of String, Integer))
                ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.6)
                CopyOverrideDirectory(
                    InstallTemp & ArchiveBaseFolder & If(OverrideHome = "." OrElse OverrideHome = "./", "", OverrideHome), '#5613
                    $"{PathMcFolder}versions\{VersionName}",
                    Task, 0.4)
            End Sub) With {
            .ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        End If
        '获取 Mod 列表
        Dim ModList As New List(Of Integer)
        Dim ModOptionalList As New List(Of Integer)
        For Each ModEntry In If(Json("files"), {})
            If ModEntry("projectID") Is Nothing OrElse ModEntry("fileID") Is Nothing Then
                Hint("某项 Mod 缺少必要信息，已跳过：" & ModEntry.ToString)
                Continue For
            End If
            ModList.Add(ModEntry("fileID"))
            If ModEntry("required") IsNot Nothing AndAlso Not ModEntry("required").ToObject(Of Boolean) Then ModOptionalList.Add(ModEntry("fileID"))
        Next
        If ModList.Any Then
            Dim ModDownloadLoaders As New List(Of LoaderBase)
            '获取 Mod 下载信息
            ModDownloadLoaders.Add(New LoaderTask(Of Integer, JArray)(GetLang("LangModModpackTaskGetModDownloadInfo"),
            Sub(Task As LoaderTask(Of Integer, JArray))
                Task.Output = GetJson(DlModRequest("https://api.curseforge.com/v1/mods/files", "POST", "{""fileIds"": [" & Join(ModList, ",") & "]}", "application/json"))("data")
                '如果文件已被删除，则 API 会跳过那一项
                If ModList.Count > Task.Output.Count Then Throw New Exception(GetLang("LangModModpackExceptionOnlineModDeleted"))
            End Sub) With {.ProgressWeight = ModList.Count / 10}) '每 10 Mod 需要 1s
            '构造 NetFile
            ModDownloadLoaders.Add(New LoaderTask(Of JArray, List(Of NetFile))(GetLang("LangModModpackTaskInitModDownloadInfo"),
            Sub(Task As LoaderTask(Of JArray, List(Of NetFile)))
                Dim FileList As New Dictionary(Of Integer, NetFile)
                For Each ModJson In Task.Input
                    Dim Id As Integer = ModJson("id").ToObject(Of Integer)
                    '跳过重复的 Mod（疑似 CurseForge Bug）
                    If FileList.ContainsKey(Id) Then Continue For
                    '可选 Mod 提示
                    If ModOptionalList.Contains(Id) Then
                        If MyMsgBox(GetLang("LangModModpackDialogContentDownloadOptionalFile", ModJson("displayName").ToString),
                                        GetLang("LangModModpackDialogTitleDownloadOptionalFile"), GetLang("LangDialogBtnYes"), GetLang("LangDialogBtnNo")) = 2 Then
                            Continue For
                        End If
                    End If
                    '根据 modules 和文件名后缀判断资源类型
                    Dim TargetFolder As String, Type As CompType
                    If ModJson("modules").Any Then 'modules 可能返回 null（#1006）
                        Dim ModuleNames = CType(ModJson("modules"), JArray).Select(Function(l) l("name").ToString).ToList
                        If ModuleNames.Contains("META-INF") OrElse ModuleNames.Contains("mcmod.info") OrElse
                           ModJson?("FileName")?.ToString?.EndsWithF(".jar", True) Then
                            TargetFolder = "mods" : Type = CompType.Mod
                        ElseIf ModuleNames.Contains("pack.mcmeta") Then
                            TargetFolder = "resourcepacks" : Type = CompType.ResourcePack
                        Else
                            TargetFolder = "shaderpacks" : Type = CompType.Shader
                        End If
                    Else
                        TargetFolder = "mods" : Type = CompType.Mod
                    End If
                    '建立 CompFile
                    Dim File As New CompFile(ModJson, Type)
                    If Not File.Available Then Continue For
                    '实际的添加
                    FileList.Add(Id, File.ToNetFile($"{PathMcFolder}versions\{VersionName}\{TargetFolder}\"))
                    Task.Progress += 1 / (1 + ModList.Count)
                Next
                Task.Output = FileList.Values.ToList
            End Sub) With {.ProgressWeight = ModList.Count / 200, .Show = False}) '每 200 Mod 需要 1s
            '下载 Mod 文件
            ModDownloadLoaders.Add(New LoaderDownload(GetLang("LangModModpackTaskDownloadMod"), New List(Of NetFile)) With {.ProgressWeight = ModList.Count * 1.5}) '每个 Mod 需要 1.5s
            '构造加载器
            InstallLoaders.Add(New LoaderCombo(Of Integer)(GetLang("LangModModpackTaskDownloadModMainLoader"), ModDownloadLoaders) With
                {.Show = False, .ProgressWeight = ModDownloadLoaders.Sum(Function(l) l.ProgressWeight)})
        End If

        '构造加载器
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = Json("minecraft")("version").ToString,
            .ForgeVersion = ForgeVersion,
            .NeoForgeVersion = NeoForgeVersion,
            .FabricVersion = FabricVersion,
            .QuiltVersion = QuiltVersion
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallModpack"), InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallGame"), MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderTask(Of String, String)(GetLang("LangModModpackTaskOrganizeFiles"),
        Sub(Task As LoaderTask(Of String, String))
            '设置图标
            Dim VersionFolder As String = $"{PathMcFolder}versions\{VersionName}\"
            If Logo IsNot Nothing AndAlso File.Exists(Logo) Then
                File.Copy(Logo, VersionFolder & "PCL\Logo.png", True)
                WriteIni(VersionFolder & "PCL\Setup.ini", "Logo", "PCL\Logo.png")
                WriteIni(VersionFolder & "PCL\Setup.ini", "LogoCustom", "True")
                Log("[ModPack] 已设置整合包 Logo：" & Logo)
            End If
            '删除原始整合包文件
            For Each Target As String In {VersionFolder & "原始整合包.zip", VersionFolder & "原始整合包.mrpack"}
                If File.Exists(Target) Then
                    Log("[ModPack] 删除原始整合包文件：" & Target)
                    File.Delete(Target)
                End If
            Next
            If File.Exists(FileAddress) AndAlso GetFileNameWithoutExtentionFromPath(FileAddress) = "modpack" Then
                Log("[ModPack] 删除安装整合包文件：" & FileAddress)
                File.Delete(FileAddress)
            End If
        End Sub) With {.ProgressWeight = 0.1, .Show = False})

        '重复任务检查
        Dim LoaderName As String = GetLang("LangModModpackTaskCurseForgeModpackInstall") & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint(GetLang("LangModModpackHintInstalling"), HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function
#End Region

#Region "Modrinth"
    Private Function InstallPackModrinth(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String, Optional VersionName As String = Nothing, Optional Logo As String = Nothing) As LoaderCombo(Of String)

        '读取 Json 文件
        Dim Json As JObject
        Try
            Json = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "modrinth.index.json").Open))
        Catch ex As Exception
            Throw New Exception(GetLang("LangModModpackExceptionModrinthModpackError"), ex)
        End Try
        If Json("dependencies") Is Nothing OrElse Json("dependencies")("minecraft") Is Nothing Then Throw New Exception(GetLang("LangModModpackModrinthNoGameInfo"))
        '获取 Mod API 版本信息
        Dim MinecraftVersion As String = Nothing
        Dim ForgeVersion As String = Nothing
        Dim NeoForgeVersion As String = Nothing
        Dim FabricVersion As String = Nothing
        Dim QuiltVersion As String = Nothing
        For Each Entry As JProperty In If(Json("dependencies"), {})
            Select Case Entry.Name.ToLower
                Case "minecraft"
                    MinecraftVersion = Entry.Value.ToString
                Case "forge" 'eg. 14.23.5.2859 / 1.19-41.1.0
                    ForgeVersion = Entry.Value.ToString
                    Log("[ModPack] 整合包 Forge 版本：" & ForgeVersion)
                Case "neoforge", "neo-forge" 'eg. 20.6.98-beta
                    NeoForgeVersion = Entry.Value.ToString
                    Log("[ModPack] 整合包 NeoForge 版本：" & NeoForgeVersion)
                Case "fabric-loader" 'eg. 0.14.14
                    FabricVersion = Entry.Value.ToString
                    Log("[ModPack] 整合包 Fabric 版本：" & FabricVersion)
                Case "quilt-loader" 'eg. 0.26.0
                    QuiltVersion = Entry.Value.ToString
                    Log("[ModPack] 整合包 Quilt 版本：" & QuiltVersion)
                Case Else
                    Hint(GetLang("LangModModpackHintUnknownModLoader", Entry.Value, Entry.Value.ToString), HintType.Critical)
                    Throw New CancelledException
            End Select
        Next
        '获取版本名
        If VersionName Is Nothing Then
            VersionName = If(Json("name"), "")
            Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
            If Validate.Validate(VersionName) <> "" Then VersionName = ""
            If VersionName = "" Then VersionName = MyMsgBoxInput(GetLang("LangModModpackInputInstanceName"), "", "", New ObjectModel.Collection(Of Validate) From {Validate})
            If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        End If
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)(GetLang("LangModModpackTaskExtraModpack"),
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.5)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "overrides",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.4)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "client-overrides",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.1)
        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '获取下载文件列表
        Dim FileList As New List(Of NetFile)
        For Each File In If(Json("files"), {})
            '检查是否需要该文件
            If File("env") IsNot Nothing Then
                Select Case File("env")("client").ToString
                    Case "optional"
                        If MyMsgBox(GetLang("LangModModpackDialogContentDownloadOptionalFile", GetFileNameFromPath(File("path").ToString)),
                                    GetLang("LangModModpackDialogTitleDownloadOptionalFile"), GetLang("LangDialogBtnYes"), GetLang("LangDialogBtnNo")) = 2 Then
                            Continue For
                        End If
                    Case "unsupported"
                        Continue For
                End Select
            End If
            '添加下载文件
            Dim Urls = File("downloads").SelectMany(Function(x) CompFile.HandleCurseForgeDownloadUrls(x.ToString())).ToList()
            Urls.AddRange(Urls.Select(Function(u) DlSourceModGet(u)).ToList)
            Urls = Urls.Distinct.ToList()
            Dim TargetPath As String = $"{PathMcFolder}versions\{VersionName}\{File("path")}"
            If Not IO.Path.GetFullPath(TargetPath).StartsWithF($"{PathMcFolder}versions\{VersionName}\", True) Then
                MyMsgBox($"整合包的文件路径超出了版本文件夹，请向整合包作者反馈此问题！" & vbCrLf & "错误的文件：" & TargetPath, "文件路径校验失败", IsWarn:=True)
                Throw New CancelledException
            End If
            FileList.Add(New NetFile(Urls, TargetPath,
                New FileChecker(ActualSize:=File("fileSize").ToObject(Of Long), Hash:=File("hashes")("sha1").ToString), True))
        Next
        If FileList.Any Then
            InstallLoaders.Add(New LoaderDownload(GetLang("LangModModpackTaskDownloadExtraFile"), FileList) With {.ProgressWeight = FileList.Count * 1.5}) '每个 Mod 需要 1.5s
        End If

        '构造加载器
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = MinecraftVersion,
            .ForgeVersion = ForgeVersion,
            .NeoForgeVersion = NeoForgeVersion,
            .FabricVersion = FabricVersion,
            .QuiltVersion = QuiltVersion
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallModpack"), InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallGame"), MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderTask(Of String, String)(GetLang("LangModModpackTaskOrganizeFiles"),
        Sub(Task As LoaderTask(Of String, String))
            '设置图标
            Dim VersionFolder As String = $"{PathMcFolder}versions\{VersionName}\"
            If Logo IsNot Nothing AndAlso File.Exists(Logo) Then
                File.Copy(Logo, VersionFolder & "PCL\Logo.png", True)
                WriteIni(VersionFolder & "PCL\Setup.ini", "Logo", "PCL\Logo.png")
                WriteIni(VersionFolder & "PCL\Setup.ini", "LogoCustom", "True")
                Log("[ModPack] 已设置整合包 Logo：" & Logo)
            End If
            '删除原始整合包文件
            For Each Target As String In {VersionFolder & "原始整合包.zip", VersionFolder & "原始整合包.mrpack"}
                If File.Exists(Target) Then
                    Log("[ModPack] 删除原始整合包文件：" & Target)
                    File.Delete(Target)
                End If
            Next
            If File.Exists(FileAddress) AndAlso GetFileNameWithoutExtentionFromPath(FileAddress) = "modpack" Then
                Log("[ModPack] 删除安装整合包文件：" & FileAddress)
                File.Delete(FileAddress)
            End If
        End Sub) With {.ProgressWeight = 0.1, .Show = False})

        '重复任务检查
        Dim LoaderName As String = GetLang("LangModModpackTaskModrinthModpackInstall") & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint(GetLang("LangModModpackHintInstalling"), HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function
#End Region

#Region "HMCL"
    Private Function InstallPackHMCL(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String) As LoaderCombo(Of String)
        '读取 Json 文件
        Dim Json As JObject
        Try
            Json = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "modpack.json").Open, Encoding.UTF8))
        Catch ex As Exception
            Throw New Exception(GetLang("LangModModpackExceptionHMCLModpackError"), ex)
        End Try
        '获取版本名
        Dim VersionName As String = If(Json("name"), "")
        Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
        If Validate.Validate(VersionName) <> "" Then VersionName = ""
        If VersionName = "" Then VersionName = MyMsgBoxInput(GetLang("LangModModpackInputInstanceName"), "", "", New ObjectModel.Collection(Of Validate) From {Validate})
        If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)(GetLang("LangModModpackTaskExtraModpack"),
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.6)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "minecraft",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.4)
        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '构造游戏本体安装加载器
        If Json("gameVersion") Is Nothing Then Throw New Exception(GetLang("LangModModpackHMCLNoGameInfo"))
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = Json("gameVersion").ToString
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase) From {
            New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallModpack"), InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)},
            New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallGame"), MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)}
        }
        '重复任务检查
        Dim LoaderName As String = GetLang("LangModModpackTaskHMCLModpackInstall") & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint(GetLang("LangModModpackHintInstalling"), HintType.Critical)
            Throw New CancelledException
        End If
        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function
#End Region

#Region "MultiMC"
    Public Class MMCPackInfo
        Public OverridedJson As New JObject
        Public AdditionalJson As New JObject
        Public Tweakers As String = Nothing
        Public JvmArgs As New JArray
        Public Libraries As New JArray
        Public IsMinecraftOverrided As Boolean = False
        Public IsForgeOverrided As Boolean = False
        Public IsNeoForgeOverrided As Boolean = False
        Public IsCleanroomOverrided As Boolean = False
        Public IsFabricOverrided As Boolean = False
        Public IsQuiltOverrided As Boolean = False
        Public IsMcArgsEdited As Boolean = False
    End Class
    Private Function InstallPackMMC(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String) As LoaderCombo(Of String)
        '读取 Json 文件
        Dim PackJson As JObject, PackInstance As String
        Dim PackInfo As MMCPackInfo = Nothing
        Try
            PackJson = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "mmc-pack.json").Open, Encoding.UTF8))
            PackInstance = ReadFile(Archive.GetEntry(ArchiveBaseFolder & "instance.cfg").Open, Encoding.UTF8)

#Region "JSON Patches"
            '参考 https://github.com/MultiMC/Launcher/wiki/JSON-Patches
            Try
                If Not Archive.Entries.Any(Function(e) e.FullName.Equals(ArchiveBaseFolder & "patches/", StringComparison.OrdinalIgnoreCase)) Then Exit Try
                Log("[ModPack] 安装的 MultiMC 整合包存在 JSON Patches")
                '排序预处理
                Dim Patches As New List(Of KeyValuePair(Of JObject, Integer))
                For Each entry In Archive.Entries
                    If Not entry.FullName.EndsWith("/") AndAlso entry.FullName.StartsWith(ArchiveBaseFolder & "patches/") Then
                        Dim Patch As JObject = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "patches/" & entry.Name).Open, Encoding.UTF8))
                        Patches.Add(New KeyValuePair(Of JObject, Integer)(Patch, If(Patch("order") IsNot Nothing, Patch("order"), 0)))
                    End If
                Next
                Dim Components As JArray = PackJson("components")
                For Each Patch In Patches
                    '检查 Patch 是否在 mmc-pack.json 中
                    Dim IsContainedInPackJson As Boolean = False
                    For Each Component In Components
                        If Component("uid").ToString() = Patch.Key("uid").ToString() Then
                            IsContainedInPackJson = True
                            Exit For
                        End If
                    Next
                    If Not IsContainedInPackJson Then
                        Log($"[ModPack] JSON-Patch {Patch.Key("uid")} 未包含于 mmc-pack.json, 跳过该 Patch")
                        Patches.Remove(Patch)
                        Continue For
                    End If
                Next
                Patches.Sort(Function(x, y) x.Value.CompareTo(y.Value))
                '应用 Patches
                PackInfo = New MMCPackInfo

                Dim Tweakers As String = Nothing
                Dim AssetIndex As JObject = Nothing
                Dim JavaVerJson As JObject = Nothing
                Dim MainClass As String = Nothing
                Dim GameArguments As New JArray
                Dim JvmArguments As New JArray
                Dim LibJson As New JArray
                Dim AddLibJson As New JArray
                For Each Patch In Patches
                    Dim PatchJson As JObject = Patch.Key
                    If PatchJson("uid") = "net.minecraft" Then
                        PackInfo.IsMinecraftOverrided = True
                    ElseIf PatchJson("uid") = "net.minecraftforge" Then
                        If PatchJson("version").ToString.StartsWithF("0.") Then
                            PackInfo.IsCleanroomOverrided = True
                        Else
                            PackInfo.IsForgeOverrided = True
                        End If
                    ElseIf PatchJson("uid") = "net.neoforged" Then
                        PackInfo.IsNeoForgeOverrided = True
                    ElseIf PatchJson("uid") = "net.fabricmc.fabric-loader" Then
                        PackInfo.IsFabricOverrided = True
                    ElseIf PatchJson("uid") = "org.quiltmc.quilt-loader" Then
                        PackInfo.IsQuiltOverrided = True
                    End If
                    'JVM 参数
                    If PatchJson("+jvmArgs") IsNot Nothing Then
                        JvmArguments.Merge(PatchJson("+jvmArgs"))
                        Log($"[ModPack] 已应用 JSON-Patch {PatchJson("uid")} 的 JVM 参数")
                    End If
                    'Libraries
                    If PatchJson("libraries") IsNot Nothing OrElse PatchJson("+libraries") IsNot Nothing Then
                        Dim Libs As New JArray
                        If PatchJson("libraries") IsNot Nothing Then
                            For Each Library In PatchJson("libraries")
                                Dim LibJobj = CType(Library, JObject)
                                If LibJobj("MMC-hint") IsNot Nothing Then
                                    LibJobj.Add("hint", LibJobj("MMC-hint"))
                                    LibJobj.Remove("MMC-hint")
                                End If
                                Libs.Add(LibJobj)
                            Next
                        End If
                        If PatchJson("+libraries") IsNot Nothing Then
                            For Each Library In PatchJson("+libraries") 'TODO: 此处处理不严谨，但也能用吧
                                Dim LibJobj = CType(Library, JObject)
                                If LibJobj("MMC-hint") IsNot Nothing Then
                                    LibJobj.Add("hint", LibJobj("MMC-hint"))
                                    LibJobj.Remove("MMC-hint")
                                End If
                                Libs.Add(LibJobj)
                            Next
                        End If
                        LibJson.Merge(Libs)
                        Log($"[ModPack] 已应用 JSON-Patch {PatchJson("uid")} 的 Libraries")
                    End If
                    'Tweakers
                    If PatchJson("+tweakers") IsNot Nothing Then
                        Tweakers = PatchJson("+tweakers")(0)
                        Log($"[ModPack] 已应用 JSON-Patch {PatchJson("uid")} 的 Tweakers")
                    End If
                    'AssetIndex
                    If PatchJson("assetIndex") IsNot Nothing Then
                        AssetIndex = PatchJson("assetIndex")
                        Log($"[ModPack] 已应用 JSON-Patch {PatchJson("uid")} 的 AssetIndex")
                    End If
                    'minecraftArguments -> arguments.game
                    If PatchJson("minecraftArguments") IsNot Nothing Then
                        For Each Arg In PatchJson("minecraftArguments").ToString().Split(" ")
                            GameArguments.Add(Arg)
                        Next
                        PackInfo.IsMcArgsEdited = True
                        Log($"[ModPack] 已应用 JSON-Patch {PatchJson("uid")} 的 minecraftArguments 至 arguments.game")
                    End If
                    'mainClass
                    If PatchJson("mainClass") IsNot Nothing Then
                        MainClass = PatchJson("mainClass")
                        Log($"[ModPack] 已应用 JSON-Patch {PatchJson("uid")} 的 mainClass")
                    End If
                    'Java 版本要求
                    If PatchJson("compatibleJavaMajors") IsNot Nothing Then
                        Dim JavaVersion As Integer = 0
                        Dim JavaComponent As String = Nothing
                        Dim JavaMajors As JArray = PatchJson("compatibleJavaMajors")
                        For Each Java In JavaMajors
                            If JavaVersion > Val(Java) Then Continue For
                            '优先选择主要的版本
                            If Val(Java) = 21 Then
                                JavaVersion = 21
                                JavaComponent = "java-runtime-delta"
                            ElseIf Val(Java) = 17 Then
                                JavaVersion = 17
                                JavaComponent = "java-runtime-gamma"
                            ElseIf Val(Java) = 11 Then
                                JavaVersion = 11
                                JavaComponent = Nothing
                            ElseIf Val(Java) = 8 Then
                                JavaVersion = 8
                                JavaComponent = "jre-legacy"
                            End If
                        Next
                        If JavaVersion = 0 Then
                            JavaVersion = JavaMajors(0)
                            JavaComponent = Nothing
                        End If
                        JavaVerJson = New JObject From {{"majorVersion", JavaVersion}}
                        If JavaComponent IsNot Nothing Then
                            JavaVerJson.Add("component", JavaComponent)
                        End If
                        Log($"[ModPack] JSON-Patch {PatchJson("uid")} 要求 Java 版本: " & JavaVersion)
                    End If
                Next
                Dim JsonArguments As JObject = Nothing
                If Not String.IsNullOrWhiteSpace(Tweakers) Then
                    GameArguments.Add("--tweakClass")
                    GameArguments.Add(Tweakers)
                End If
                If GameArguments IsNot Nothing OrElse JvmArguments IsNot Nothing Then
                    JvmArguments.Insert(0, "-Djava.library.path=${natives_directory}")
                    JvmArguments.Insert(1, "-Dminecraft.launcher.brand=${launcher_name}")
                    JvmArguments.Insert(2, "-Dminecraft.launcher.version=${launcher_version}")
                    JvmArguments.Insert(3, "-cp")
                    JvmArguments.Insert(4, "${classpath}")
                    JsonArguments = New JObject From {
                        {"game", GameArguments},
                        {"jvm", JvmArguments}
                    }
                End If
                PackInfo.OverridedJson = New JObject
                If JsonArguments IsNot Nothing Then PackInfo.OverridedJson.Add("arguments", JsonArguments)
                If MainClass IsNot Nothing Then PackInfo.OverridedJson.Add("mainClass", MainClass)
                If AssetIndex IsNot Nothing Then PackInfo.OverridedJson.Add("assetIndex", AssetIndex)
                If JavaVerJson IsNot Nothing Then PackInfo.OverridedJson.Add("javaVersion", JavaVerJson)
                If LibJson IsNot Nothing Then PackInfo.OverridedJson.Add("libraries", LibJson)
            Catch ex As Exception
                Log(ex, $"应用 MMC JSON-Patches 失败")
            End Try
#End Region

        Catch ex As Exception
            Throw New Exception(GetLang("LangModModpackExceptionMMCModpackError"), ex)
        End Try
        '获取版本名
        Dim VersionName As String = If(RegexSeek(PackInstance, "(?<=\nname\=)[^\n]+"), "")
        Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
        If Validate.Validate(VersionName) <> "" Then VersionName = ""
        If VersionName = "" Then VersionName = MyMsgBoxInput(GetLang("LangModModpackInputInstanceName"), "", "", New ObjectModel.Collection(Of Validate) From {Validate})
        If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim SetupFile As String = $"{PathMcFolder}versions\{VersionName}\PCL\Setup.ini"
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)(GetLang("LangModModpackTaskExtraModpack"),
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.55)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "libraries",
                PathMcFolder & "versions\" & VersionName & "\libraries",
                Task, 0.2)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & ".minecraft",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.2)

#Region "instance.cfg"
            '读取 MMC 设置文件（#2655）
            Try
                Dim MMCSetupFile As String = InstallTemp & ArchiveBaseFolder & "instance.cfg"
                If File.Exists(MMCSetupFile) Then
                    '将其中的等号替换为冒号，以符合 ini 文件格式
                    Dim Lines As New List(Of String)
                    For Each Line In ReadFile(MMCSetupFile).Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        If Not Line.Contains("=") Then Continue For
                        Lines.Add(Line.BeforeFirst("=") & ":" & Line.AfterFirst("="))
                    Next
                    WriteFile(MMCSetupFile, Join(Lines, vbCrLf))
                    '读取文件
                    If ReadIni(MMCSetupFile, "OverrideCommands", False) Then
                        Dim PreLaunchCommand As String = ReadIni(MMCSetupFile, "PreLaunchCommand")
                        If PreLaunchCommand <> "" Then
                            PreLaunchCommand = PreLaunchCommand.Replace("\""", """").
                                Replace("$INST_JAVA", "{java}javaw.exe").
                                Replace("$INST_MC_DIR\", "{minecraft}").Replace("$INST_MC_DIR", "{minecraft}").
                                Replace("$INST_DIR\", "{verpath}").Replace("$INST_DIR", "{verpath}").
                                Replace("$INST_ID", "{name}").Replace("$INST_NAME", "{name}")
                            WriteIni(SetupFile, "VersionAdvanceRun", PreLaunchCommand)
                            Log("[ModPack] 迁移 MultiMC 版本独立设置：启动前执行命令：" & PreLaunchCommand)
                        End If
                    End If
                    If ReadIni(MMCSetupFile, "JoinServerOnLaunch", False) Then
                        Dim ServerAddress As String = ReadIni(MMCSetupFile, "JoinServerOnLaunchAddress").Replace("\""", """")
                        WriteIni(SetupFile, "VersionServerEnter", ServerAddress)
                        Log("[ModPack] 迁移 MultiMC 版本独立设置：自动进入服务器：" & ServerAddress)
                    End If
                    If ReadIni(MMCSetupFile, "IgnoreJavaCompatibility", False) Then
                        WriteIni(SetupFile, "VersionAdvanceJava", True)
                        Log("[ModPack] 迁移 MultiMC 版本独立设置：忽略 Java 兼容性警告")
                    End If
                    Dim Logo As String = ReadIni(MMCSetupFile, "iconKey", "")
                    If Logo <> "" AndAlso File.Exists($"{InstallTemp}{ArchiveBaseFolder}{Logo}.png") Then
                        WriteIni(SetupFile, "LogoCustom", True)
                        WriteIni(SetupFile, "Logo", "PCL\Logo.png")
                        CopyFile($"{InstallTemp}{ArchiveBaseFolder}{Logo}.png", $"{PathMcFolder}versions\{VersionName}\PCL\Logo.png")
                        Log($"[ModPack] 迁移 MultiMC 版本独立设置：版本图标（{Logo}.png）")
                    End If
                    'JVM 参数
                    Dim JvmArgs As String = ReadIni(MMCSetupFile, "JvmArgs", "")
                    If JvmArgs <> "" Then
                        If ReadIni(MMCSetupFile, "OverrideJavaArgs", False) Then
                            WriteIni(SetupFile, "VersionAdvanceJvm", JvmArgs)
                            Log("[ModPack] 迁移 MultiMC 版本独立设置：JVM 参数（覆盖）：" & JvmArgs)
                        Else
                            JvmArgs += " " & Setup.Get("LaunchAdvanceJvm")
                            WriteIni(SetupFile, "VersionAdvanceJvm", JvmArgs)
                            Log("[ModPack] 迁移 MultiMC 版本独立设置：JVM 参数（追加）：" & JvmArgs)
                        End If
                    End If
                End If
            Catch ex As Exception
                Log(ex, $"读取 MMC 配置文件失败（{InstallTemp}{ArchiveBaseFolder}instance.cfg）")
            End Try
#End Region

        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '构造版本安装请求
        If PackJson("components") Is Nothing Then Throw New Exception(GetLang("LangModModpackMMCNoGameInfo"))
        Dim Request As New McInstallRequest With {.TargetVersionName = VersionName, .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\"}
        For Each Component In PackJson("components")
            Select Case If(Component("uid"), "").ToString
                Case "org.lwjgl"
                    Log("[ModPack] 已跳过 LWJGL 项")
                Case "net.minecraft"
                    Request.MinecraftName = Component("version")
                Case "net.minecraftforge"
                    If Component("version").ToString().StartsWithF("0.") Then
                        Request.CleanroomVersion = Component("version")
                    Else
                        Request.ForgeVersion = Component("version")
                    End If
                Case "net.neoforged"
                    Request.NeoForgeVersion = Component("version")
                Case "net.fabricmc.fabric-loader"
                    Request.FabricVersion = Component("version")
                Case "org.quiltmc.quilt-loader"
                    Request.QuiltVersion = Component("version")
            End Select
        Next
        If PackInfo IsNot Nothing Then Request.MMCPackInfo = PackInfo
        '构造加载器
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallModpack"), InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallGame"), MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})

        '重复任务检查
        Dim LoaderName As String = GetLang("LangModModpackTaskMMCModpackInstall") & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint(GetLang("LangModModpackHintInstalling"), HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function
#End Region

#Region "MCBBS"
    Private Function InstallPackMCBBS(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String,
                                      Optional VersionName As String = Nothing) As LoaderCombo(Of String)
        '读取 Json 文件
        Dim Json As JObject
        Try
            Dim Entry = If(Archive.GetEntry(ArchiveBaseFolder & "mcbbs.packmeta"), Archive.GetEntry(ArchiveBaseFolder & "manifest.json"))
            Json = GetJson(ReadFile(Entry.Open, Encoding.UTF8))
        Catch ex As Exception
            Throw New Exception(GetLang("LangModModpackExceptionMCBBSModpackError"), ex)
        End Try
        '获取版本名
        If VersionName Is Nothing Then
            VersionName = If(Json("name"), "")
            Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
            If Validate.Validate(VersionName) <> "" Then VersionName = ""
            If VersionName = "" Then VersionName = MyMsgBoxInput(GetLang("LangModModpackInputInstanceName"), "", "", New ObjectModel.Collection(Of Validate) From {Validate})
            If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        End If
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim SetupFile As String = $"{PathMcFolder}versions\{VersionName}\PCL\Setup.ini"
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)(GetLang("LangModModpackTaskExtraModpack"),
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.6)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "overrides",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.4)
            'JVM 参数
            If Json("launchInfo") IsNot Nothing Then
                Dim LaunchInfo As JObject = Json("launchInfo")
                If LaunchInfo.ContainsKey("javaArgument") Then WriteIni(SetupFile, "VersionAdvanceJvm", String.Join(" ", LaunchInfo("javaArgument")))
                If LaunchInfo.ContainsKey("launchArgument") Then WriteIni(SetupFile, "VersionAdvanceGame", String.Join(" ", LaunchInfo("launchArgument")))
            End If
        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '构造加载器
        If Json("addons") Is Nothing Then Throw New Exception(GetLang("LangModModpackMCBBSNoGameInfoAddons"))
        Dim Addons As New Dictionary(Of String, String)
        For Each Entry In Json("addons")
            Addons.Add(Entry("id"), Entry("version"))
        Next
        If Not Addons.ContainsKey("game") Then
            Hint(GetLang("LangModModpackMCBBSNoGameInfo"), HintType.Critical)
            Return Nothing
        End If
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = Addons("game"),
            .OptiFineVersion = If(Addons.ContainsKey("optifine"), Addons("optifine"), Nothing),
            .ForgeVersion = If(Addons.ContainsKey("forge"), Addons("forge"), Nothing),
            .NeoForgeVersion = If(Addons.ContainsKey("neoforge"), Addons("neoforge"), Nothing),
            .FabricVersion = If(Addons.ContainsKey("fabric"), Addons("fabric"), Nothing),
            .QuiltVersion = If(Addons.ContainsKey("quilt"), Addons("quilt"), Nothing)
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallModpack"), InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)(GetLang("LangModModpackTaskInstallGame"), MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})

        '重复任务检查
        Dim LoaderName As String = GetLang("LangModModpackTaskMCBBSModpackInstall") & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint(GetLang("LangModModpackHintInstalling"), HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        'If Archive IsNot Nothing Then Archive.Dispose() '解除占用，以免在加载器中触发 “正由另一进程使用，因此该进程无法访问此文件”
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function
#End Region

#Region "带启动器的压缩包"
    Private Function InstallPackLauncherPack(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String) As LoaderCombo(Of String)
        '获取解压路径
        MyMsgBox(GetLang("LangModModpackDialogContentInstallTip"), GetLang("LangModModpackDialogTitleInstall"), GetLang("LangDialogBtnContinue"), ForceWait:=True)
        Dim TargetFolder As String = SelectFolder(GetLang("LangModModpackSelectInstallFolder"))
        If String.IsNullOrEmpty(TargetFolder) Then Throw New CancelledException
        If Directory.GetFileSystemEntries(TargetFolder).Length > 0 Then Hint(GetLang("LangModModpackFolderShouldBeEmpty"), HintType.Critical) : Throw New CancelledException
        '解压
        Dim Loader As New LoaderCombo(Of String)(GetLang("LangModModpackTaskExtractArchive"), {
            New LoaderTask(Of String, Integer)(GetLang("LangModModpackTaskExtractArchive"),
            Sub(Task As LoaderTask(Of String, Integer))
                ExtractModpackFiles(TargetFolder, FileAddress, Task, 0.9)
                Thread.Sleep(400) '避免文件争用
                '查找解压后的 exe 文件
                Dim Launcher As String = Nothing
                For Each ExeFile In Directory.GetFiles(TargetFolder, "*.exe", SearchOption.TopDirectoryOnly)
                    Dim Info = FileVersionInfo.GetVersionInfo(ExeFile)
                    Log($"[Modpack] 文件 {ExeFile} 的产品名标识为 {Info.ProductName}")
                    If Info.ProductName = "Plain Craft Launcher" Then
                        Launcher = ExeFile
                        Log($"[Modpack] 发现整合包附带的 PCL 启动器：{ExeFile}")
                    ElseIf (Info.ProductName.ContainsF("Launcher", True) OrElse Info.ProductName.ContainsF("启动", True)) AndAlso
                        Not Info.ProductName = "Plain Craft Launcher Admin Manager" Then
                        If Launcher Is Nothing Then
                            Launcher = ExeFile
                            Log($"[Modpack] 发现整合包附带的疑似第三方启动器：{ExeFile}")
                        End If
                    End If
                Next
                Task.Progress = 0.95
                '尝试使用附带的启动器打开
                If Launcher IsNot Nothing Then
                    Log("[Modpack] 找到压缩包中附带的启动器：" & Launcher)
                    If MyMsgBox(GetLang("LangModModpackDialogContentAnotherLauncher", Launcher), GetLang("LangModModpackDialogTitleAnotherLauncher"), GetLang("LangModModpackDialogBtnAnotherLauncher1"), GetLang("LangModModpackDialogBtnAnotherLauncher2")) = 1 Then
                        OpenExplorer(TargetFolder)
                        ShellOnly(Launcher, "--wait") '要求等待已有的 PCL 退出
                        Log("[Modpack] 为换用整合包中的启动器启动，强制结束程序")
                        FrmMain.EndProgram(False)
                        Return
                    End If
                Else
                    Log("[Modpack] 未找到压缩包中附带的启动器")
                End If
                OpenExplorer(TargetFolder)
                '加入文件夹列表
                Dim VersionName As String = GetFolderNameFromPath(TargetFolder)
                Directory.CreateDirectory(TargetFolder & ".minecraft\")
                PageSelectLeft.AddFolder(
                    TargetFolder & ".minecraft\" & ArchiveBaseFolder.Replace("/", "\").TrimStart("\"), '格式例如：包裹文件夹\.minecraft\（最短为空字符串）
                    VersionName, False)
                '调用 modpack 文件进行安装
                Dim ModpackFile = Directory.GetFiles(TargetFolder, "modpack.*", SearchOption.AllDirectories).First
                Log("[Modpack] 调用 modpack 文件继续安装：" & ModpackFile)
                ModpackInstall(ModpackFile)
            End Sub)
        })
        Loader.Start(TargetFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        FrmMain.BtnExtraDownload.Ribble()
        Return Loader
    End Function
#End Region

#Region "普通压缩包"
    Private Function InstallPackCompress(FileAddress As String, Archive As Compression.ZipArchive) As LoaderCombo(Of String)
        '尝试定位 .minecraft 文件夹：寻找形如 “/versions/XXX/XXX.json” 的路径
        Dim Match As RegularExpressions.Match = Nothing
        Dim Regex As New RegularExpressions.Regex("^.*\/(?=versions\/(?<ver>[^\/]+)\/(\k<ver>)\.json$)", RegularExpressions.RegexOptions.IgnoreCase)
        For Each Entry In Archive.Entries
            Dim EntryMatch = Regex.Match("/" & Entry.FullName)
            If EntryMatch.Success Then
                Match = EntryMatch
                Exit For
            End If
        Next
        If Match Is Nothing Then Throw New Exception(GetLang("LangModModpackExceptionNoModpackFound")) '没有匹配
        Dim ArchiveBaseFolder As String = Match.Value.Replace("/", "\").TrimStart("\") '格式例如：包裹文件夹\.minecraft\（最短为空字符串）
        Dim VersionName As String = Match.Groups(1).Value
        Log("[ModPack] 检测到压缩包的 .minecraft 根目录：" & ArchiveBaseFolder & "，命中的版本名：" & VersionName)
        '获取解压路径
        MyMsgBox(GetLang("LangModModpackDialogContentInstallTip"), GetLang("LangModModpackDialogTitleInstall"), GetLang("LangDialogBtnContinue"), ForceWait:=True)
        Dim TargetFolder As String = SelectFolder(GetLang("LangModModpackSelectInstallFolder"))
        If String.IsNullOrEmpty(TargetFolder) Then Throw New CancelledException
        If TargetFolder.Contains("!") OrElse TargetFolder.Contains(";") Then Hint(GetLang("LangModModpackFolderNoExclamationOrSemicolon"), HintType.Critical) : Throw New CancelledException
        If Directory.GetFileSystemEntries(TargetFolder).Length > 0 Then Hint(GetLang("LangModModpackFolderShouldBeEmpty"), HintType.Critical) : Throw New CancelledException
        '解压
        Dim Loader As New LoaderCombo(Of String)(GetLang("LangModModpackTaskExtractArchive"), {
            New LoaderTask(Of String, Integer)(GetLang("LangModModpackTaskExtractArchive"),
            Sub(Task As LoaderTask(Of String, Integer))
                ExtractModpackFiles(TargetFolder, FileAddress, Task, 0.95)
                '加入文件夹列表
                PageSelectLeft.AddFolder(TargetFolder & ArchiveBaseFolder, GetFolderNameFromPath(TargetFolder), False)
                Thread.Sleep(400) '避免文件争用
                RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.VersionSelect))
            End Sub)
        }) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(TargetFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        FrmMain.BtnExtraDownload.Ribble()
        Return Loader
    End Function

#End Region

End Module
