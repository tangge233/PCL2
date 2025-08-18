﻿Imports System.Collections.ObjectModel

Class PageLoginProfile
    ''' <summary>
    ''' 刷新页面显示的所有信息。
    ''' </summary>
    Public Sub Reload() Handles Me.Loaded
        RefreshProfileList()
        FrmLoginProfileSkin = Nothing
        RunInNewThread(Sub()
                           Thread.Sleep(800)
                           RunInUi(Sub() FrmLaunchLeft.RefreshPage(True))
                       End Sub)
    End Sub
    Public Property ProfileCollection As New ObservableCollection(Of ProfileItem)
    Public Class ProfileItem
        Public ReadOnly Property Info As String
        Public ReadOnly Property Logo As String
        Public ReadOnly Property Profile As McProfile
        Public ReadOnly Property Username As String
            Get
                Return Profile.Username
            End Get
        End Property
        Public Sub New(profile As McProfile)
            Me.Profile = profile
            Info = GetProfileInfo(profile)
            Dim LogoPath As String = PathTemp & $"Cache\Skin\Head\{Profile.SkinHeadId}.png"
            If Not (File.Exists(LogoPath) AndAlso Not New FileInfo(LogoPath).Length = 0) Then
                LogoPath = ModBase.Logo.IconButtonUser
            End If
            Logo = LogoPath
        End Sub
    End Class
    ''' <summary>
    ''' 刷新档案列表
    ''' </summary>
    Public Sub RefreshProfileList()
        Log("[Profile] 刷新档案列表")
        ProfileCollection.Clear()
        GetProfile()
        Try
            For Each Profile In ProfileList
            ProfileCollection.Add(New ProfileItem(Profile))
            Next
            Log($"[Profile] 档案列表刷新完成")
        Catch ex As Exception
            Log(ex, "读取档案列表失败", LogLevel.Feedback)
        End Try
        If Not ProfileList.Any() Then
            Setup.Set("HintProfileSelect", True)
            HintCreate.Visibility = Visibility.Visible
        Else
            HintCreate.Visibility = Visibility.Collapsed
        End If
    End Sub

#Region "控件"
    Private Sub SelectProfile(sender As Object, e As MouseButtonEventArgs)
        SelectedProfile = CType(sender, MyListItem).Tag
        Log($"[Profile] 选定档案: {sender.Tag.Username}, 以 {sender.Tag.Type} 方式验证")
        LastUsedProfile = ProfileList.IndexOf(sender.Tag) '获取当前档案的序号
        RunInUi(Sub()
                    FrmLaunchLeft.RefreshPage(True)
                    FrmLaunchLeft.BtnLaunch.IsEnabled = True
                End Sub)
    End Sub
    Private Sub ProfileContMenuBuild(sender As MyListItem, e As EventArgs)
        Dim BtnUUID As New MyIconButton With {.Logo = Logo.IconButtonInfo, .ToolTip = "更改 UUID", .Tag = sender.Tag}
        ToolTipService.SetPlacement(BtnUUID, Primitives.PlacementMode.Center)
        ToolTipService.SetVerticalOffset(BtnUUID, 30)
        ToolTipService.SetHorizontalOffset(BtnUUID, 2)
        AddHandler BtnUUID.Click, AddressOf EditProfileUuid
        Dim BtnServerName As New MyIconButton With {.Logo = Logo.IconButtonInfo, .ToolTip = "更改验证服务器名称", .Tag = sender.Tag}
        ToolTipService.SetPlacement(BtnServerName, Primitives.PlacementMode.Center)
        ToolTipService.SetVerticalOffset(BtnServerName, 30)
        ToolTipService.SetHorizontalOffset(BtnServerName, 2)
        AddHandler BtnServerName.Click, AddressOf EditProfileServer
        Dim BtnDelete As New MyIconButton With {.Logo = Logo.IconButtonDelete, .ToolTip = "删除档案", .Tag = sender.Tag}
        ToolTipService.SetPlacement(BtnDelete, Primitives.PlacementMode.Center)
        ToolTipService.SetVerticalOffset(BtnDelete, 30)
        ToolTipService.SetHorizontalOffset(BtnDelete, 2)
        AddHandler BtnDelete.Click, AddressOf DeleteProfile
        If sender.Tag.Type = 0 Then
            sender.Buttons = {BtnUUID, BtnDelete}
        ElseIf sender.Tag.Type = 3 Then
            sender.Buttons = {BtnDelete}
        Else
            sender.Buttons = {BtnDelete}
        End If
    End Sub
    '创建档案
    Private Sub BtnNew_Click(sender As Object, e As EventArgs) Handles BtnNew.Click
        RunInNewThread(Sub()
                           CreateProfile()
                           RunInUi(Sub() RefreshProfileList())
                       End Sub)
    End Sub
    '编辑 UUID
    Private Sub EditProfileUuid(sender As Object, e As EventArgs)
        EditOfflineUuid(sender.Tag)
    End Sub
    '编辑验证服务器名称
    Private Sub EditProfileServer(sender As Object, e As EventArgs)
        Dim Name As String = MyMsgBoxInput("修改验证服务器名称", $"请输入新的验证服务器名称", sender.Tag.ServerName)
        If Name IsNot Nothing Then
            EditAuthServerName(sender.Tag, Name)
        End If
    End Sub
    '删除档案
    Private Sub DeleteProfile(sender As Object, e As EventArgs)
        If MyMsgBox($"你正在选择删除此档案，该操作无法撤销。{vbCrLf}确定继续？", "删除档案确认", "继续", "取消", IsWarn:=True, ForceWait:=True) = 2 Then Exit Sub
        RemoveProfile(sender.Tag)
        RunInUi(Sub() RefreshProfileList())
    End Sub
    '导入 / 导出档案
    Private Sub BtnPort_Click() Handles BtnPort.Click
        MigrateProfile()
        RunInUi(Sub() RefreshProfileList())
    End Sub
#End Region

End Class
