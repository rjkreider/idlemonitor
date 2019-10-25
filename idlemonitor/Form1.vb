﻿Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1

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
        'Public dwTime As Int32
        Public dwTime As Integer
    End Structure

    Private objExitWin As New cWrapExitWindows()

    Dim Count As Integer = 0
    Dim Count2 As Integer = 0
    Dim count3 As Integer = 0
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label4.Text = Environment.UserName.ToString
        Label1.Text = 0
        Label2.Text = 0

        Timer1.Interval = 1000
        Timer2.Interval = 1000
        Timer1.Start()
        Timer2.Start()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Dim LastInput As New tagLASTINPUTINFO()
        Dim IdleTime As Int32
        LastInput.cbSize = CUInt(Marshal.SizeOf(LastInput))
        LastInput.dwTime = 0

        If GetLastInputInfo(LastInput) Then
            IdleTime = System.Environment.TickCount - LastInput.dwTime
            Label1.Text = Math.Round(IdleTime / 1000, 0)
        End If

        Label2.Text = CStr(Count)
        Count = Count + 1
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        'If CInt(Label1.Text) < 500 Then
        '    Count = 0
        'End If
        'If Count = 300 Then 'five minutes
        '    Timer2.Stop()
        '    ' objExitWin.ExitWindows(cWrapExitWindows.Action.LogOff)
        'End If

        If CInt(Label1.Text) < 5 Then

            count3 = Count2 + count3

            Count2 = 0

        Else
            Count2 += 1



        End If
        Application.DoEvents()
        ' DO WRITE LOG EVENTS
        If CInt(Label1.Text) > 5 Then
            Label3.Text = "Status:  IDLE (" & Count2 & ") " & GetCaptionOfActiveWindow()
        Else
            If CInt(Label1.Text) < 5 Then
                Label3.Text = "Status ACTIVE Total Idle: " & count3
            End If
        End If


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