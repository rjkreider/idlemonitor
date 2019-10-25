Imports System.Runtime.InteropServices
Friend Class InputSource
    <DllImport("User32.dll")>
    Private Shared Function GetLastInputInfo(ByRef plii As LASTINPUTINFO) As Boolean
    End Function

    <DllImport("Kernel32.dll")>
    Private Shared Function GetLastError() As UInteger
    End Function

    Public Function GetLastInputTime() As Date
        Dim lastInputInfo = New LASTINPUTINFO
        lastInputInfo.dwTime = 0
        lastInputInfo.cbSize = CUInt(Marshal.SizeOf(lastInputInfo))
        GetLastInputInfo(lastInputInfo)
        Return Date.Now.AddMilliseconds(-(Environment.TickCount - lastInputInfo.dwTime))
    End Function

    Public Shared Function GetIdleTime() As UInteger
        Dim lastinput = New LASTINPUTINFO
        lastinput.cbSize = CUInt(Marshal.SizeOf(lastinput))
        GetLastInputInfo(lastinput)
        Return CUInt(Environment.TickCount) - lastinput.dwTime
    End Function

    Public Shared Function GetTickCount() As Long
        Return Environment.TickCount
    End Function

    Public Shared Function GetLastInput() As Long
        Dim lastInPut = New LASTINPUTINFO
        lastInPut.cbSize = CUInt(Marshal.SizeOf(lastInPut))

        If Not GetLastInputInfo(lastInPut) Then
            Throw New Exception(GetLastError.ToString)
        End If

        Return lastInPut.dwTime
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure LASTINPUTINFO
        Public cbSize As UInteger
        Public dwTime As UInteger
    End Structure
End Class
