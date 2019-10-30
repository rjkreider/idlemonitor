Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Public Class IdleMonitor

    <DllImport("user32.dll")>
    Public Shared Function GetLastInputInfo(ByRef plii As tagLASTINPUTINFO) As [Boolean]
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetWindowText(hWnd As IntPtr, text As StringBuilder, count As Integer) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
    End Function

    Public Structure tagLASTINPUTINFO
        Public cbSize As UInteger
        Public dwTime As Integer
    End Structure

    Private objExitWin As New cWrapExitWindows()

    Dim idleStartDelay As Integer = 300 ' Seconds idle before considering user as idle
    Dim timeElapsed As Integer = 0 ' Total seconds elapsed
    Dim currIdlePeriod As Integer = 0 ' How many seconds idle since hitting idleStartDelay
    Dim totalIdleTime As Integer = 0 ' Total of all idle time after idleStartDelay
    'TODO:
    Dim currIdleTime As Integer = 0 ' How many seconds idle since hitting idleStartDelay (this can be refactored and just use currIdlePeriod-idleStartDelay)
    Dim activeWindow As String = Nothing
    Dim logStart As String = Nothing
    Dim logEnd As String = Nothing
    Dim useBase64 As Boolean = True ' Base64 each event when logging to file
    Dim logFile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\" & Environment.UserName & ".log"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Visible = False
        Me.Hide()
        Me.ShowInTaskbar = False
        NotifyIcon1.Visible = True
        ShowInTaskbar = False
        Timer1.Interval = 1000
        Timer2.Interval = 1000
        Timer1.Start()
        Timer2.Start()
    End Sub

    Private Sub things_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.L And Control.ModifierKeys = (Keys.Control + Keys.Shift + Keys.Alt) Then
            readLog()
        End If
        If e.KeyCode = Keys.X And Control.ModifierKeys = (Keys.Control + Keys.Shift + Keys.Alt) Then
            Application.Exit()
        End If
        If e.KeyCode = Keys.O And Control.ModifierKeys = (Keys.Control + Keys.Shift + Keys.Alt) Then
            logviewer.Show()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Dim LastInput As New tagLASTINPUTINFO()
        Dim IdleTime As Int32
        LastInput.cbSize = CUInt(Marshal.SizeOf(LastInput))
        LastInput.dwTime = 0

        If GetLastInputInfo(LastInput) Then
            IdleTime = System.Environment.TickCount - LastInput.dwTime
            currIdleTime = Math.Round(IdleTime / 1000, 0)
        End If

        timeElapsed += 1

    End Sub
    Private Sub readLog()
        TextBox1.Visible = True
        TextBox1.Clear()
        Timer1.Stop()
        Timer2.Stop()
        Dim b As Byte()
        For Each line As String In File.ReadAllLines(logFile)
            b = System.Convert.FromBase64String(line)
            TextBox1.AppendText(System.Text.ASCIIEncoding.ASCII.GetString(b) & vbCrLf)
        Next

        Timer1.Start()
        Timer2.Start()

    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        If currIdleTime >= idleStartDelay Then
            currIdlePeriod += 1
            activeWindow = GetCaptionOfActiveWindow()
            If currIdlePeriod = 1 Then
                logStart = Now() & "," & Environment.UserName & "," & activeWindow
            End If

        Else

            totalIdleTime += currIdlePeriod
            If currIdlePeriod > 0 Then
                logEnd = currIdlePeriod
                logIdle(logStart, logEnd)
            End If

            currIdleTime = 0
            currIdlePeriod = 0
        End If

    End Sub

    Private Sub logIdle(ByVal action As String, ByVal data As String)
        Dim str As String = action & "," & data
        Dim base64Encoded As String = Nothing

        Dim logStr As String = Nothing
        If useBase64 Then
            base64Encoded = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes(str))
            logStr = base64Encoded
        Else
            logStr = str
        End If
        logToFile(logStr)
    End Sub

    Private Sub logToFile(ByVal str As String)
        Dim fileExists As Boolean = File.Exists(logFile)
        Using sw As New StreamWriter(File.Open(logFile, FileMode.Append))
            sw.WriteLine(
                IIf(fileExists, str, str))
        End Using
    End Sub
    Private Function GetCaptionOfActiveWindow() As String
        Dim strTitle As String = String.Empty
        Dim handle As IntPtr = GetForegroundWindow()
        ' Obtain the length of the text   
        Dim intLength As Integer = GetWindowTextLength(handle) + 1
        Dim stringBuilder As New StringBuilder(intLength)
        If GetWindowText(handle, stringBuilder, intLength) > 0 Then
            strTitle = stringBuilder.ToString()
        End If
        Return strTitle
    End Function

    Private Sub IdleMonitor_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize, MyBase.Closing
        ' Don't exit application, minimize to systemtray.
        If Me.WindowState = FormWindowState.Minimized Then
            NotifyIcon1.Visible = True
            ShowInTaskbar = False
            TextBox1.Clear()
            TextBox1.Visible = False
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ' Don't exit application, minimize to systemtray.
        Me.WindowState = FormWindowState.Minimized
        e.Cancel = True
    End Sub

    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        Me.Show()
        TextBox1.Visible = False
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub

    Private Function rot13(ByVal str As String) As StringBuilder
        ' Not using this anymore, but here JIC
        Dim result As StringBuilder = New StringBuilder()
        For Each ch As Char In str

            If (Not Char.IsLetter(ch)) Then
                result.Append(ch)
                Continue For
            End If

            Dim checkIndex As Integer = Asc("a") - (Char.IsUpper(ch) * -32)
            Dim index As Integer = ((Asc(ch) - checkIndex) + 13) Mod 26

            result.Append(Chr(index + checkIndex))

        Next
        Return result
    End Function





End Class

Public Class cWrapExitWindows

    Private Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Int32, ByVal dwReserved As Int32) As Boolean
    Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As IntPtr
    Private Declare Sub OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As Int32, ByRef TokenHandle As IntPtr)
    Private Declare Sub LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As Long)
    Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As IntPtr, ByVal DisableAllPrivileges As Boolean, ByRef NewState As LUID, ByVal BufferLength As Int32, ByVal PreviousState As IntPtr, ByVal ReturnLength As IntPtr) As Boolean

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Friend Structure LUID
        Public Count As Integer
        Public LUID As Long
        Public Attribute As Integer
    End Structure

    Public Enum Action
        LogOff = 0
        Shutdown = 1
        Restart = 2
        PowerOff = 8
    End Enum

    Public Sub ExitWindows(ByVal how As Action, Optional ByVal Forced As Boolean = True)
        Dim TokenPrivilege As LUID
        Dim hProcess As IntPtr = GetCurrentProcess()
        Dim hToken As IntPtr = IntPtr.Zero
        OpenProcessToken(hProcess, &H28, hToken)
        TokenPrivilege.Count = 1
        TokenPrivilege.LUID = 0
        TokenPrivilege.Attribute = 2
        LookupPrivilegeValue(Nothing, "SeShutdownPrivilege", TokenPrivilege.LUID)
        AdjustTokenPrivileges(hToken, False, TokenPrivilege, 0, IntPtr.Zero, IntPtr.Zero)
        If Forced Then
            ExitWindowsEx(how + 4, 0)
        Else
            ExitWindowsEx(how, 0)
        End If
    End Sub

End Class