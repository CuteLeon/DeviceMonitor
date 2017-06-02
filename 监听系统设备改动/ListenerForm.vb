Imports System.ComponentModel

Public Class ListenerForm
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewinteger As Integer) As Integer
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Integer, ByVal hwnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As DEV_BROADCAST_HDR, ByRef Source As Integer, ByVal Length As Integer) As Integer
    Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As DEV_BROADCAST_VOLUME, ByRef Source As Integer, ByVal Length As Integer) As Integer
    Private Delegate Function SubWndProc(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
    Private Const GWL_WNDPROC = (-4)
    Private OldWinProc As Integer
    Private Const WM_PAINT = &HF
    Private Const WM_DEVICECHANGE = &H219
    Private Const DBT_DEVICEARRIVAL = &H8000
    Private Const DBT_DEVICEREMOVECOMPLETE = &H8004
    Private Const DBT_DEVTYP_VOLUME = &H2
    Private Structure DEV_BROADCAST_HDR
        Dim lSize As Integer
        Dim lDevicetype As Integer
        Dim lReserved As Integer
    End Structure

    Private Structure DEV_BROADCAST_VOLUME
        Dim lSize As Integer
        Dim lDevicetype As Integer
        Dim lReserved As Integer
        Dim lUnitMask As Integer
        Dim iFlag As Int16
    End Structure

    Private info As DEV_BROADCAST_HDR
    Private info_volume As DEV_BROADCAST_VOLUME

    Private Sub ListenerForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim leftWndProc As SubWndProc = AddressOf MyWndproc
        OldWinProc = GetWindowLong(Me.Handle, GWL_WNDPROC) '记录原始的委托地址，方便退出程序时恢复
        SetWindowLong(Me.Handle, GWL_WNDPROC, Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(leftWndProc)) '开始截取消息
        Me.Text = "初始委托地址：&H" & Hex(OldWinProc)
        If OldWinProc = 0 Then MsgBox("注册钩子失败！") : End
    End Sub

    Private Sub ListenerForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        SetWindowLong(Me.Handle, GWL_WNDPROC, OldWinProc) '取消消息截取
    End Sub

    Private Function MyWndproc(ByVal hwnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        'Debug.Print(Msg & vbCrLf & wParam & vbCrLf & lParam)
        If Msg = WM_DEVICECHANGE Then
            If wParam = DBT_DEVICEARRIVAL Then '插入设备
                CopyMemory(info, lParam, Len(info))
                Debug.Print("挂载设备！")
                If info.lDevicetype = DBT_DEVTYP_VOLUME Then
                    CopyMemory(info_volume, lParam, Len(info_volume))
                    MsgBox(Chr(GetDriveName(info_volume.lUnitMask)),, "盘符：")
                End If
            ElseIf wParam = DBT_DEVICEREMOVECOMPLETE Then             '移走USB
                Debug.Print("移除设备！")
                CopyMemory(info, lParam, Len(info))
                If info.lDevicetype = DBT_DEVTYP_VOLUME Then
                    CopyMemory(info_volume, lParam, Len(info_volume))
                End If
            End If
        End If

        Return CallWindowProc(OldWinProc, hwnd, Msg, wParam, lParam)
    End Function

    Private Function GetDriveName(ByVal lUnitMask As Long) As Byte
        Dim i As Long
        i = 0

        Do While lUnitMask Mod 2 <> 1
            lUnitMask = lUnitMask \ 2
            i = i + 1
        Loop

        GetDriveName = Asc("A") + i
    End Function

End Class
