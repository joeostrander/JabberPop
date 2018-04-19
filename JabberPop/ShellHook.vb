

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Imports System.Runtime.InteropServices
Imports System.Windows.Forms


Public Class ShellHook

    'Public Shared hwnd As IntPtr

    '<DllImport("user32.dll", CharSet:=CharSet.Unicode)> _
    'Public Shared Function FindWindow(sClassName As IntPtr, sAppName As [String]) As IntPtr
    'End Function

    '<DllImport("kernel32.dll")> _
    'Public Shared Function GetLastError() As UInteger
    'End Function

    'Public Enum ShellEvents As Integer
    '    HSHELL_WINDOWCREATED = 1
    '    HSHELL_WINDOWDESTROYED = 2
    '    HSHELL_ACTIVATESHELLWINDOW = 3
    '    HSHELL_WINDOWACTIVATED = 4
    '    HSHELL_GETMINRECT = 5
    '    HSHELL_REDRAW = 6
    '    HSHELL_TASKMAN = 7
    '    HSHELL_LANGUAGE = 8
    '    HSHELL_ACCESSIBILITYSTATE = 11
    '    HSHELL_APPCOMMAND = 12
    '    HSHELL_FLASH = 32774
    'End Enum

    'Private msgNotify As Integer
    'Public Delegate Sub EventHandler(sender As Object, data As String)
    'Public Event WindowEvent As EventHandler
    'Protected Overridable Sub OnWindowEvent(data As String)
    '    Dim handler = WindowEvent
    '    RaiseEvent handler(Me, data)
    'End Sub

    'Public Declare Ansi Function RegisterWindowMessage Lib "user3.dll" Alias "RegisterWindowMessageA" (lpString As String) As Integer

    'Public Declare Ansi Function DeregisterShellHookWindow Lib "user32" (hWnd As IntPtr) As Integer

    'Public Declare Ansi Function RegisterShellHookWindow Lib "user32" (hWnd As IntPtr) As Integer

    'Public Declare Ansi Function GetWindowText Lib "user32" Alias "GetWindowTextA" (hwnd As IntPtr, lpString As System.Text.StringBuilder, cch As Integer) As Integer

    'Public Declare Ansi Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (hwnd As IntPtr) As Integer

    'Public Sub New()
    '    ' Hook on to the shell
    '    msgNotify = RegisterWindowMessage("SHELLHOOK")
    '    RegisterShellHookWindow(Me.Handle)
    'End Sub

    'Protected Overrides Sub WndProc(ByRef m As Message)

    '    If m.Msg = msgNotify Then
    '        ' Receive shell messages
    '        Select Case CType(m.WParam.ToInt32(), ShellEvents)
    '            Case ShellEvents.HSHELL_WINDOWCREATED
    '                Exit Select
    '            Case ShellEvents.HSHELL_WINDOWDESTROYED
    '                Exit Select
    '            Case ShellEvents.HSHELL_WINDOWACTIVATED
    '                Exit Select
    '            Case ShellEvents.HSHELL_FLASH
    '                Dim wName As String = GetWindowName(m.LParam)
    '                Dim action = CType(m.WParam.ToInt32(), ShellEvents)
    '                OnWindowEvent(String.Format("{0} - {1}: {2}", action, m.LParam, wName))
    '                Exit Select
    '        End Select
    '    End If

    '    MyBase.WndProc(m)
    'End Sub

    'Private Function GetWindowName(hwnd As IntPtr) As String
    '    Dim sb As New StringBuilder()
    '    Dim longi As Integer = GetWindowTextLength(hwnd) + 1
    '    sb.Capacity = longi
    '    GetWindowText(hwnd, sb, sb.Capacity)
    '    Return sb.ToString()
    'End Function

    'Protected Overrides Sub Dispose(disposing As Boolean)
    '    Try
    '        DeregisterShellHookWindow(Me.Handle)
    '    Catch
    '    End Try
    '    MyBase.Dispose(disposing)
    'End Sub

    ''Private Shared Sub Main(args As String())
    ''    'hwnd = FindWindow(IntPtr.Zero, args[0]);
    ''    hwnd = FindWindow(IntPtr.Zero, "Remote PC Admin")
    ''    Console.WriteLine(String.Format("Window Handle:  {0}", hwnd.ToInt32().ToString("X")))

    ''    Dim f = New ShellHook()
    ''    AddHandler f.WindowEvent, Function(sender, data) Console.WriteLine(data)
    ''    While True
    ''        Application.DoEvents()
    ''    End While


    ''End Sub
End Class

