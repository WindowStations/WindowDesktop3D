Public Class Form1
    Public Sub New()
        On Error Resume Next
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        If Process.GetProcessesByName("WindowContextMenu").Length > 1 Then
            For Each p As Process In Process.GetProcessesByName("WindowContextMenu")
                If p.Id <> Process.GetCurrentProcess.Id Then
                    p.Kill()
                End If
            Next
        End If
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Dim scm As New ShellContext.ShellContextMenu
        Dim pth() As String
        Dim files(0) As System.IO.FileInfo
        Me.Size = New Size(1, 1)
        Me.Location = New Point(0, 0)
        pth = Environment.GetCommandLineArgs()
        files(0) = New System.IO.FileInfo(pth(1))
        scm.ShowContextMenu(files, MousePosition)
        Timer1.Enabled = True
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        On Error Resume Next
        If Me.Visible = True Then
            Me.Visible = False
            Timer1.Interval = 2000
        End If
        If Process.GetProcessesByName("Window3D").Length = 0 Then
            Me.Close()
        End If
        'If Process.GetProcessesByName("WindowLauncher").Length = 0 Then
        '    Me.Close()
        'End If
        'If Process.GetProcessesByName("WindowStations").Length = 0 Then
        '    Me.Close()
        'End If
    End Sub
End Class
