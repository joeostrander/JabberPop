Imports System.Text
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class Form1


    Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal _
                lpPoint As Point) As IntPtr
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure 'RECT
    Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Int32, ByRef lpRect As RECT) As Boolean

    Public Enum ShowWindowCommands As Integer
        ''' <summary>
        ''' Hides the window and activates another window.
        ''' </summary>
        Hide = 0
        ''' <summary>
        ''' Activates and displays a window. If the window is minimized or 
        ''' maximized, the system restores it to its original size and position.
        ''' An application should specify this flag when displaying the window 
        ''' for the first time.
        ''' </summary>
        Normal = 1
        ''' <summary>
        ''' Activates the window and displays it as a minimized window.
        ''' </summary>
        ShowMinimized = 2
        ''' <summary>
        ''' Maximizes the specified window.
        ''' </summary>
        Maximize = 3
        ' is this the right value?
        ''' <summary>
        ''' Activates the window and displays it as a maximized window.
        ''' </summary>       
        ShowMaximized = 3
        ''' <summary>
        ''' Displays a window in its most recent size and position. This value 
        ''' is similar to <see cref="Win32.ShowWindowCommands.Normal"/>, except 
        ''' the window is not actived.
        ''' </summary>
        ShowNoActivate = 4
        ''' <summary>
        ''' Activates the window and displays it in its current size and position. 
        ''' </summary>
        Show = 5
        ''' <summary>
        ''' Minimizes the specified window and activates the next top-level 
        ''' window in the Z order.
        ''' </summary>
        Minimize = 6
        ''' <summary>
        ''' Displays the window as a minimized window. This value is similar to
        ''' <see cref="Win32.ShowWindowCommands.ShowMinimized"/>, except the 
        ''' window is not activated.
        ''' </summary>
        ShowMinNoActive = 7
        ''' <summary>
        ''' Displays the window in its current size and position. This value is 
        ''' similar to <see cref="Win32.ShowWindowCommands.Show"/>, except the 
        ''' window is not activated.
        ''' </summary>
        ShowNA = 8
        ''' <summary>
        ''' Activates and displays the window. If the window is minimized or 
        ''' maximized, the system restores it to its original size and position. 
        ''' An application should specify this flag when restoring a minimized window.
        ''' </summary>
        Restore = 9
        ''' <summary>
        ''' Sets the show state based on the SW_* value specified in the 
        ''' STARTUPINFO structure passed to the CreateProcess function by the 
        ''' program that started the application.
        ''' </summary>
        ShowDefault = 10
        ''' <summary>
        '''  <b>Windows 2000/XP:</b> Minimizes a window, even if the thread 
        ''' that owns the window is not responding. This flag should only be 
        ''' used when minimizing windows from a different thread.
        ''' </summary>
        ForceMinimize = 11
    End Enum

    Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Int32, ByVal nCmdShow As ShowWindowCommands) As Boolean

    Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpwindowtext As System.Text.StringBuilder, ByVal nmaxcount As Int32) As Int32
    Private Declare Auto Function GetWindowTextLength Lib "user32.dll" (ByVal hwnd As IntPtr) As Int32

    Private Declare Ansi Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Private Declare Ansi Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer

    Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
    Private Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
    Private Declare Function EnumChildWindows Lib "user32.dll" _
        (ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess, _
        ByVal lParam As IntPtr) As Boolean

    Public Declare Auto Function GetClassName Lib "User32.dll" (ByVal hwnd As IntPtr, <Out()> ByVal lpClassName As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer

    'Public Declare Function ObjectFromLresult Lib "Oleacc.dll" (ByVal lResult As Integer, ByRef riid As Guid, ByVal wParam As Integer, ByRef ppvObject As mshtml.HTMLDocument) As Integer
    Public Declare Ansi Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (lpString As String) As Integer
    Public Declare Ansi Function RegisterShellHookWindow Lib "user32.dll" (hWnd As IntPtr) As Integer

    'Public Declare Function SendMessageTimeoutA Lib "User32.dll" (ByVal hwnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer, ByVal fuFlags As Integer, ByVal uTimeout As Integer, ByRef lpdwResult As Integer) As Integer

    Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As IntPtr, ByVal hWnd2 As IntPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr

    Private windowList As List(Of IntPtr)
    Public Delegate Function MyDelegateCallBack(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean
    Private Declare Function EnumWindows Lib "user32" (ByVal x As MyDelegateCallBack, ByVal y As Integer) As Integer

    Private msgNotify As Integer
    'Public Delegate Sub EventHandler(sender As Object, data As String)
    'Public Event WindowEvent As EventHandler
    'Protected Overridable Sub OnWindowEvent(data As String)
    '    RaiseEvent WindowEvent(Me, data)
    'End Sub

    Public Enum ShellEvents As Integer
        HSHELL_WINDOWCREATED = 1
        HSHELL_WINDOWDESTROYED = 2
        HSHELL_ACTIVATESHELLWINDOW = 3
        HSHELL_WINDOWACTIVATED = 4
        HSHELL_GETMINRECT = 5
        HSHELL_REDRAW = 6
        HSHELL_TASKMAN = 7
        HSHELL_LANGUAGE = 8
        HSHELL_ACCESSIBILITYSTATE = 11
        HSHELL_APPCOMMAND = 12
        HSHELL_FLASH = 32774
    End Enum


    Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As UInteger)

    Const HWND_TOPMOST As Integer = -1
    Const HWND_NOTOPMOST As Integer = -2
    Const SWP_NOSIZE As Integer = 1

    Private LastShake As DateTime = Now.AddDays(-1)



    Private boolStartup As Boolean = False


    Private Function EnumWindowProc(ByVal hwnd As IntPtr, ByVal lParam As Integer) As Boolean
        Try
            If IsWindowVisible(hwnd) Then
                Dim sb As StringBuilder
                sb = New StringBuilder(128)
                GetClassName(hwnd, sb, sb.Capacity)
                Dim strWindowClass As String = sb.ToString
                If strWindowClass = "wcl_manager1" Then
                    windowList.Add(hwnd)
                End If
            End If
        Catch ex As Exception
            EnumWindowProc = False
            Exit Function
        End Try

        EnumWindowProc = True
    End Function

    Private Function GetJabberWindow() As IntPtr
        'If the Chat conversations window only has 1 chat, it will not be titled "Conversations
        'In that case... we'll need to check each wcl_manager1 window for existence of 2 child "Internet Explorer_Server" windows

        Dim hWnd_Jabber As IntPtr = IntPtr.Zero
        Dim hWnd_In As IntPtr = IntPtr.Zero
        Dim hWnd_Out As IntPtr = IntPtr.Zero



        windowList = New List(Of IntPtr)

        hWnd_Jabber = FindWindow("wcl_manager1", "Conversations")

        If IsWindow(hWnd_Jabber) Then
            windowList.Add(hWnd_Jabber)
        Else

            Try
                Dim del As MyDelegateCallBack
                del = New MyDelegateCallBack(AddressOf EnumWindowProc)
                EnumWindows(del, 0)
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try


        End If

        For Each hWnd As IntPtr In windowList
            Dim intChildCount As Integer = 0

            For Each chld As IntPtr In GetChildWindows(hWnd)
                Dim sb As StringBuilder
                sb = New StringBuilder(128)
                GetClassName(chld, sb, sb.Capacity)
                Dim strWindowClass As String = sb.ToString
                If strWindowClass = "ATL:RICHEDIT50W" Then
                    'outgoing
                    hWnd_Out = chld
                End If
                If strWindowClass = "Internet Explorer_Server" Then
                    'incoming
                    hWnd_In = chld
                End If
            Next
            If hWnd_In <> IntPtr.Zero AndAlso hWnd_Out <> IntPtr.Zero Then
                hWnd_Jabber = hWnd
                Exit For
            End If
        Next

        Return hWnd_Jabber


    End Function


    Private Function GetWindowTextByHandle(ByVal hWnd As IntPtr) As String
        Dim len As Integer = GetWindowTextLength(hWnd)
        Dim strText As String = ""
        If len > 0 Then
            Dim b As New System.Text.StringBuilder(ChrW(0), len + 1)
            Dim ret = GetWindowText(hWnd, b, b.Capacity)
            If ret <> 0 Then strText = b.ToString
        End If
        GetWindowTextByHandle = strText
    End Function


    Private Function EnumWindow(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
        Dim ChildrenList As List(Of IntPtr) = GCHandle.FromIntPtr(Parameter).Target
        If ChildrenList Is Nothing Then Throw New Exception("GCHandle Target could not be cast as List(Of IntPtr)")
        ChildrenList.Add(Handle)
        Return True
    End Function

    Private Function GetChildWindows(ByVal ParentHandle As IntPtr) As IntPtr()
        Dim ChildrenList As New List(Of IntPtr)
        Dim ListHandle As GCHandle = GCHandle.Alloc(ChildrenList)
        Try
            EnumChildWindows(ParentHandle, AddressOf EnumWindow, GCHandle.ToIntPtr(ListHandle))
        Finally
            If ListHandle.IsAllocated Then ListHandle.Free()
        End Try
        Return ChildrenList.ToArray
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        ' Hook on to the shell
        Try
            msgNotify = RegisterWindowMessage("SHELLHOOK")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        RegisterShellHookWindow(Me.Handle)

        LoadRegistrySettings()
    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        NotifyIcon1.Text = Application.ProductName
        Me.Text = Application.ProductName

        If Me.WindowState = FormWindowState.Minimized Then Me.Hide()

    End Sub

    Private Sub Form1_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
        Else
            'ShowMe()
        End If
    End Sub


    'Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    If boolTrayExit = False Then
    '        If MsgBox("Are you sure you want to exit?" & vbCrLf & vbCrLf & "(Click No to minimize to the Notification area)", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Exit " & Application.ProductName & "?") <> MsgBoxResult.Yes Then
    '            e.Cancel = True
    '            Me.WindowState = FormWindowState.Minimized
    '            Exit Sub
    '        End If
    '    End If
    'End Sub

    Private Sub LoadRegistrySettings()
        On Error Resume Next

        Dim regKey_RUN As RegistryKey
        regKey_RUN = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True)

        Dim appRegKey As RegistryKey
        appRegKey = Registry.CurrentUser.OpenSubKey("Software\" & Application.ProductName, True)

        If IsNothing(appRegKey) Then
            'Create the key
            Registry.CurrentUser.CreateSubKey("Software\" & Application.ProductName)
            appRegKey = Registry.CurrentUser.OpenSubKey("Software\" & Application.ProductName, True)
            If IsNothing(appRegKey) Then
                regKey_RUN.Close()
                Exit Sub
            End If

            'Ask user if they want it to launch auto...
            If MsgBox("Launch " & Application.ProductName & " at Startup?", _
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                            Application.ProductName) = MsgBoxResult.Yes Then
                boolStartup = True
            Else
                boolStartup = False
            End If

        Else
            'Load current Startup Value and set it in the interface
            boolStartup = appRegKey.GetValue("RunAtStartup", False)

        End If

        regKey_RUN.Close()
        appRegKey.Close()

        SaveRegistrySettings()

    End Sub

    Private Sub SaveRegistrySettings()
        On Error Resume Next

        Dim regKey_RUN As RegistryKey
        regKey_RUN = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True)

        Dim appRegKey As RegistryKey
        appRegKey = Registry.CurrentUser.OpenSubKey("Software\" & Application.ProductName, True)

        'Startup 
        appRegKey.SetValue("RunAtStartup", boolStartup, RegistryValueKind.DWord)

        'If boolStartup=true, add executable path to the Run registry key
        If boolStartup = True Then
            regKey_RUN.SetValue(Application.ProductName, Application.ExecutablePath)
        Else
            regKey_RUN.DeleteValue(Application.ProductName)
        End If

        RunAtStartupToolStripMenuItemTRAY.Checked = boolStartup


        regKey_RUN.Close()
        appRegKey.Close()

    End Sub



    Protected Overrides Sub WndProc(ByRef m As Message)

        If m.Msg = msgNotify Then

            ' Receive shell messages
            Select Case CType(m.WParam.ToInt32(), ShellEvents)
                Case ShellEvents.HSHELL_WINDOWCREATED

                Case ShellEvents.HSHELL_WINDOWDESTROYED

                Case ShellEvents.HSHELL_WINDOWACTIVATED

                Case ShellEvents.HSHELL_FLASH
                    Dim hWnd_Jabber As IntPtr = GetJabberWindow()
                    If IsWindow(hWnd_Jabber) Then
                        Dim wName As String = GetWindowTextByHandle(m.LParam)
                        Dim action = CType(m.WParam.ToInt32(), ShellEvents)
                        If hWnd_Jabber = m.LParam Then

                            'OnWindowEvent(String.Format("{0} - {1}: {2}", action, m.LParam, wName))
                            'Debug.WriteLine(String.Format("{0} - {1}: {2}", action, m.LParam, wName))
                            Dim ts As TimeSpan = Now.Subtract(LastShake)
                            If ts.TotalSeconds >= 5 Then
                                shakeMe(hWnd_Jabber)
                                LastShake = Now
                            End If

                        End If

                    End If
            End Select
        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If EnabledToolStripMenuItem.CheckState <> CheckState.Checked Then Exit Sub

        Dim hWnd As IntPtr = GetJabberWindow()
        Label1.Text = String.Format("Jabber Handle:  {0}", "0x" & Hex(hWnd.ToInt32).ToString.PadLeft(8, "0"))

    End Sub


    Public Sub shakeMe(ByRef hWnd As IntPtr)
        If EnabledToolStripMenuItem.CheckState <> CheckState.Checked Then Exit Sub
        ShowWindow(hWnd, ShowWindowCommands.Normal)

        Dim rect As New RECT
        GetWindowRect(hWnd.ToInt32, rect)

        'Dim myWidth As Integer = rect.right - rect.left
        'Dim myHeight As Integer = rect.bottom - rect.top

        Dim myLoc As Point
        Dim myLocDef As Point

        myLoc = New Point(rect.left, rect.top)

        If myLoc.X < 0 Or myLoc.Y < 0 Then
            '    myLoc.X = (Screen.PrimaryScreen.Bounds.Width / 2) - ((rect.right - rect.left) / 2)
            '    myLoc.Y = (Screen.PrimaryScreen.Bounds.Height / 2) - ((rect.bottom - rect.top) / 2) - 300
            '    '            MsgBox("Width:  " & Screen.PrimaryScreen.Bounds.Width & vbCrLf & "Height:  " & Screen.PrimaryScreen.Bounds.Height & vbCrLf & "X:  " & myLoc.X & vbCrLf & "Y:  " & myLoc.Y)
            'MsgBox(rect.right - rect.left & vbCrLf & rect.bottom - rect.top)
        End If

        myLocDef = myLoc

        For i As Integer = 0 To 10
            For x As Integer = 0 To 4
                Select Case x
                    Case 0
                        myLoc.X = myLocDef.X + 10
                    Case 1
                        myLoc.X = myLocDef.X - 10
                    Case 2
                        myLoc.Y = myLocDef.Y - 10
                    Case 3
                        myLoc.Y = myLocDef.Y + 10
                    Case 4
                        myLoc = myLocDef
                End Select
                'MoveWindow(hWnd, myLoc.X, myLoc.Y, myWidth, myHeight, False)
                SetWindowPos(hWnd, HWND_TOPMOST, myLoc.X, myLoc.Y, 0, 0, SWP_NOSIZE)
                System.Threading.Thread.Sleep(15)
            Next
        Next
        'MoveWindow(hWnd, myLocDef.X, myLocDef.Y, myWidth, myHeight, False)
        SetWindowPos(hWnd, HWND_NOTOPMOST, myLocDef.X, myLocDef.Y, 0, 0, SWP_NOSIZE)

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub RunAtStartupToolStripMenuItemTRAY_Click(sender As Object, e As EventArgs) Handles RunAtStartupToolStripMenuItemTRAY.Click
        boolStartup = RunAtStartupToolStripMenuItemTRAY.Checked
        SaveRegistrySettings()
    End Sub

  
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        'boolTrayExit = True
        Application.Exit()
    End Sub
End Class
