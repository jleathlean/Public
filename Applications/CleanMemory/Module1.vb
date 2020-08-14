Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Reflection

Namespace FreeStandbyMemory
    Class Program
        Const SE_INCREASE_QUOTA_PRIVILEGE As UInteger = &H5
        Const SE_PROF_SINGLE_PROCESS_PRIVILEGE As UInteger = &HD
        Const SystemFileCacheInformation As Integer = &H15
        Const SystemMemoryListInformation As Integer = &H50
        Shared MemoryPurgeStandbyList As Integer = &H4
        Shared retv As Boolean = False

        Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As ULong, ByVal Enable As Boolean, ByVal CurrentThread As Boolean, ByRef RetValue As Boolean) As UInteger

        Declare Function NtSetSystemInformation Lib "ntdll.dll" (ByVal InfoClass As Integer, ByRef Info As Integer, ByVal Length As Integer) As UInteger

        Public Declare Function EmptyWorkingSet Lib "psapi.dll" (hwProc As IntPtr) As Int32

        Declare Function SetSystemFileCacheSize Lib "kernel32.dll" (ByVal MinimumFileCacheSize As IntPtr, ByVal MaximumFileCacheSize As IntPtr, ByVal Flags As Integer) As Boolean

        '       <MarshalAs(UnmanagedType.Bool)>
        Private Declare Function GlobalMemoryStatusEx Lib "kernel32" Alias "GlobalMemoryStatusEx" (<[In](), Out()> ByVal lpBuffer As MEMORYSTATUSEX) As <MarshalAs(UnmanagedType.Bool)> Boolean

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
        Private Class MEMORYSTATUSEX
            Public dwLength As UInteger
            Public dwMemoryLoad As UInteger
            Public ullTotalPhys As ULong
            Public ullAvailPhys As ULong
            Public ullTotalPageFile As ULong
            Public ullAvailPageFile As ULong
            Public ullTotalVirtual As ULong
            Public ullAvailVirtual As ULong
            Public ullAvailExtendedVirtual As ULong

            Public Sub New()
                Me.dwLength = CUInt(Marshal.SizeOf(GetType(MEMORYSTATUSEX)))
            End Sub
        End Class

        Public Shared Sub Main(ByVal args As String())

            ' Based on code from: https://pastebin.com/Kj36ug5h

            Console.WriteLine("Cleaning Memory...")
            Dim availphysmem As ULong = If((args.Length = 0), UInt64.MaxValue, Convert.ToUInt64(args(0)) * 1024 * 1024)
            Dim systemcachews As Boolean = (args.Length = 0 OrElse args.Length >= 2 AndAlso args(1) = "1")
            Dim memStatus As MEMORYSTATUSEX = New MEMORYSTATUSEX()

            If GlobalMemoryStatusEx(memStatus) AndAlso memStatus.ullAvailPhys > availphysmem Then Return
            RtlAdjustPrivilege(SE_INCREASE_QUOTA_PRIVILEGE, True, False, retv)
            RtlAdjustPrivilege(SE_PROF_SINGLE_PROCESS_PRIVILEGE, True, False, retv)
            NtSetSystemInformation(SystemMemoryListInformation, MemoryPurgeStandbyList, Marshal.SizeOf(MemoryPurgeStandbyList))

            If systemcachews Then
                SetSystemFileCacheSize(IntPtr.Subtract(IntPtr.Zero, 1), IntPtr.Subtract(IntPtr.Zero, 1), 0)
                Dim processlist As Process() = Process.GetProcesses()

                For Each p As Process In processlist
                    If p.SessionId = 0 Then
                        Try
                            EmptyWorkingSet(p.Handle)
                        Catch __unusedException1__ As Exception
                        End Try
                    End If
                Next

            End If
        End Sub
    End Class
End Namespace
