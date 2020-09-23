Attribute VB_Name = "SubClassHook"
Option Explicit
'submitted by Tom Pydeski
'module for subclassing the DeviceChange Windows Message
'Windows sends all top-level windows a set of default WM_DEVICECHANGE messages when
'new devices or media (such as a CD or DVD) are added and become available, and when
'existing devices or media are removed.
'You do not need to register the application to receive these default messages
'We could do this to receive further information on devices
'
'This program should detect the following:
'-CD or DVD inserted into drive
'-Floppy inserted
'-USB thumbdrive or external hard drive added
'-USB serial port added
'-probably a lot of other hardware changes, but I could not test a lot of others
'
'Below is from http://msdn.microsoft.com/library/default.asp?url=/library/en-us/devio/base/device_events.asp
'Applications, including services, can register to receive notification of device events.
' For example, a catalog service can receive notice of volumes being mounted or
' dismounted so it can adjust the paths to files on the volume. The system notifies an
' application that a device event has occurred by sending the application a
' WM_DEVICECHANGE <wm_devicechange.asp> message. The system notifies a
' service that a device event has occurred by invoking the service's event handler
' function, HandlerEx </library/en-us/dllproc/base/handlerex.asp>.
'To receive device event notices, call the RegisterDeviceNotification
' <registerdevicenotification.asp> function with a DEV_BROADCAST_HANDLE
' <dev_broadcast_handle_str.asp> structure. Be sure to set the dbch_handle member
' to the device handle obtained from the CreateFile </library/en
'-us/fileio/fs/createfile.asp> function. Also, set the dbch_devicetype member to
' DBT_DEVTYP_HANDLE. The function returns a device notification handle. Note
' that this is not the same as the volume handle.

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public PrevProc As Long
Public m_hWnd As Long
Global CurrentVolume As Integer
Private Const DBTF_MEDIA As Long = &H1& ' Media comings and goings
'Change affects media in drive. If not set, change affects physical device or drive.
Private Const DBTF_NET As Long = &H2& 'Indicated logical volume is a network volume.
Private Const DBT_DEVTYP_OEM As Long = &H0&             'OEM- or IHV-defined device type. This structure is a DEV_BROADCAST_OEM structure.
Private Const DBT_DEVTYP_VOLUME As Long = &H2&          'Logical volume. This structure is a DEV_BROADCAST_VOLUME structure.
Private Const DBT_DEVTYP_PORT As Long = &H3&            'Port device (serial or parallel). This structure is a DEV_BROADCAST_PORT structure.
Public Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5& 'Class of devices.This structure is a DEV_BROADCAST_DEVICEINTERFACE structure.
Private Const DBT_DEVTYP_HANDLE As Long = &H6&          'File system handle. This structure is a DEV_BROADCAST_HANDLE structure.
'
Private Const DBT_DEVNODES_CHANGED As Long = &H7
Private Const DBT_DEVICEARRIVAL As Long = &H8000&
Private Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&
Private Const DBT_QUERYCHANGECONFIG = &H17&
Private Const DBT_CONFIGCHANGED = &H18&
Private Const DBT_CONFIGCHANGECANCELED = &H19&
'private Const DBT_DeviceQUERYREMOVE = &H800&
Private Const DBT_DeviceQUERYREMOVE = &H8001&
Private Const DBT_DeviceQUERYREMOVEFAILED = &H8002&
Private Const DBT_DeviceREMOVEPENDING = &H8003&
Private Const DBT_DeviceTYPESPECIFIC = &H8005&
Private Const DBT_USERDEFINED = &HFFFF&
Private Type DEV_BROADCAST_HDR
    dbch_size As Long
    dbch_devicetype As Long
    dbch_reserved As Long
End Type
Private Type DEV_BROADCAST_VOLUME
    dbcv_size As Long
    dbcv_devicetype As Long
    dbcv_reserved As Long
    dbcv_unitmask As Long
    dbcv_flags As Long
End Type
Private Type DEV_BROADCAST_PORT
    dbcp_size As Long 'Size of this structure, in bytes.
    dbcp_devicetype As Long 'Set to DBT_DEVTYP_PORT.
    dbcp_reserved As Long 'Reserved; do not use.
    dbcp_name As Long 'Pointer to a null-terminated string specifying the friendly
    'name of the port or the device connected to the port.
    'Friendly names are intended to help the user quickly
    'and accurately identify the device—for example,
    '"COM1" and "Standard 28800 bps Modem" are considered friendly names.
End Type
Public Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Public Type DEV_BROADCAST_DEVICEINTERFACE
    dbcc_size As Long 'Size of this structure, in bytes.  This is the size of the
    'members plus the actual length of the dbcc_name string
    '(the null character is accounted for by the declaration
    'of dbcc_name as a one-character array.)
    dbcc_devicetype As Long 'Set to DBT_DEVTYP_DEVICEINTERFACE.
    dbcc_reserved As Long 'Reserved; do not use.
    dbcc_classguid As Guid 'GUID for the interface device class.
    'dbcc_name As String 'A null-terminated string that specifies the name of the device.
    dbcc_name As Long 'pointer to the string
End Type
Private Type DEV_BROADCAST_HANDLE
    dbch_size  As Long 'Size of this structure, in bytes.
    dbch_devicetype  As Long 'Set to DBT_DEVTYP_HANDLE.
    dbch_reserved  As Long 'Reserved; do not use.
    dbch_handle  As Long 'Handle to the device to be checked.
    dbch_hdevnotify As Long 'Handle to the device notification. This handle is returned by RegisterDeviceNotification <registerdevicenotification.asp>.
    dbch_eventguid As Guid 'GUID for the custom event.
    dbch_nameoffset As Long 'Offset of an optional string buffer. Valid only for DBT_CUSTOMEVENT.
    dbch_data As Byte 'Optional binary data. Valid only for DBT_CUSTOMEVENT.
End Type
Dim PortName As String
'below found in an example 1/29/2007
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_NO_ROOT_DIR As Long = 1
'
Private Const DRIVE_UNKNOWN As Long = 0  'The drive type cannot be determined.
Private Const DRIVE_ABSENT As Long = 1  'The root path is invalid; for example, there is no volume is mounted at the path.
Private Const DRIVE_REMOVABLE As Long = 2  'The drive has removable media; for example, a floppy drive or flash card reader.
Private Const DRIVE_FIXED As Long = 3  'The drive has fixed media; for example, a hard drive, flash drive, or thumb drive.
Private Const DRIVE_REMOTE As Long = 4  'The drive is a remote (network) drive.
Private Const DRIVE_CDROM As Long = 5  'The drive is a CD-ROM drive.
Private Const DRIVE_RAMDISK As Long = 6  'The drive is a RAM disk.
'
'docking info can be found at
'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\IDConfigDB\CurrentDockInfo
'
'below from the example I found for DeviceNotification
'for reading the GUID
Private Declare Function StringFromGUID2 Lib "OLE32.dll" (ByRef rGUID As Any, ByVal lpsz As String, ByVal cchMax As Long) As Long
Private Declare Function lstrcpyA Lib "Kernel32.dll" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlenA Lib "Kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub GetDWord Lib "MSVBVM60.dll" Alias "GetMem4" (ByRef inSrc As Any, ByRef inDst As Long)
Private Declare Sub GetWord Lib "MSVBVM60.dll" Alias "GetMem2" (ByRef inSrc As Any, ByRef inDst As Integer)
Private Const WM_DEVICECHANGE As Long = &H219
'
Public Declare Function RegisterDeviceNotification Lib "User32.dll" Alias "RegisterDeviceNotificationA" (ByVal hRecipient As Long, ByRef NotificationFilter As Any, ByVal Flags As Long) As Long
Public Declare Function UnregisterDeviceNotification Lib "User32.dll" (ByVal Handle As Long) As Long
Public hDevNotify As Long
Public Const DEVICE_NOTIFY_WINDOW_HANDLE As Long = &H0
Public Const DEVICE_NOTIFY_ALL_INTERFACE_CLASSES As Long = &H4
Private Declare Sub CopyMemoryDBDevInterface Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEV_BROADCAST_DEVICEINTERFACE, ByVal Source As Long, ByVal Length As Long)
'for string pointer manipulation
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
Private Declare Function lstrcpyToBuffer Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
' Pointer validation in StringFromPointer
Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
'our information variables
Dim ChangeType$

Public Sub HookWin(hWnd As Long)
m_hWnd = hWnd
PrevProc = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WinProc)
End Sub

Public Sub UnHookWin()
SetWindowLong m_hWnd, GWL_WNDPROC, PrevProc
End Sub

Public Function WinProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Info$
On Error Resume Next
'Debug.Print WinMess(uMsg)
'Form1.Text1.Text = Form1.Text1.Text & WinMess(uMsg) & vbCrLf
If uMsg = WM_DEVICECHANGE Then
    'we added a new Device
    If wParam = DBT_DEVICEARRIVAL Then
        ChangeType$ = "Device added"
    ElseIf wParam = DBT_DEVICEREMOVECOMPLETE Then
        ChangeType$ = "Device removed"
    End If
    Select Case wParam
        'wParam represents the Event that has occurred.
        'This parameter can be one of the following values from the Dbt.h header file
        Case DBT_DEVICEARRIVAL
            'New Device was added or media was inserted into a drive.
            Info$ = "DBT_DeviceArrival - lParam = " & lParam & "(" & Hex$(lParam) & "H)"
        Case DBT_DEVICEREMOVECOMPLETE
            'Device was removed or media was removed from a drive.
            Info$ = "DBT_DeviceRemoveComplete - lParam = " & lParam & "(" & Hex$(lParam) & "H)"
        Case DBT_DEVNODES_CHANGED
            'A device has been added to or removed from the system.
            'The system broadcasts the DBT_DEVNODES_CHANGED device event when a device
            'has been added to or removed from the system.
            Info$ = "DBT_DEVNODES_CHANGED : lParam = " & lParam & "(" & Hex$(lParam) & "H)"
        Case Else
            'some other dbt event, which we don't care about in this example
            Info$ = "Unknown Event - wParam = " & wParam & "(" & Hex$(wParam) & "H) : " & "lParam = " & lParam & "(" & Hex$(lParam) & "H)"
    End Select
    'lparam is a Pointer to a structure that contains event-specific data.
    'Its format depends on the value of the wParam parameter.
    ProcessDeviceChange wParam, lParam
End If
If Len(Info$) > 0 Then
    Debug.Print Info$
    Form1.Text1.SelText = Info$ & vbCrLf
    'MsgBox Info$, vbExclamation, "WM_DeviceCHANGE"
End If
WinProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
End Function

Sub ProcessDeviceChange(wParamIn As Long, lParamIn As Long)
Dim i As Integer
Dim Info$
Dim DBHdr As DEV_BROADCAST_HDR
Dim DBVol As DEV_BROADCAST_VOLUME
Dim DBPort As DEV_BROADCAST_PORT
Dim DBInter As DEV_BROADCAST_DEVICEINTERFACE
Dim DBHandle As DEV_BROADCAST_HANDLE
'
Dim DevBroadcastHeader As DEV_BROADCAST_HDR
Dim UnitMask As Long
Dim Flags As Integer
Dim DeviceGUID As Guid
Dim DeviceNamePtr As Long
Dim DriveLetters As String
Dim LoopDrives As Long

'external hard drive gives lparam = 1310180(13FDE4H), which gives no flags
'
'Each WM_DEVICECHANGE message has an associated event that describes the change,
'and a structure that provides detailed information about the change.
'The structure consists of an event-independent header, DEV_BROADCAST_HDR,
'followed by event-dependent members.
'The event-dependent members describe the device to which the event applies.
'To use this structure, applications must first determine the event type
'and the device type. Then, they can use the correct structure to take appropriate action.
If lParamIn <> 0 Then
    Debug.Print "Copying Memory... ";
    'copy the lparam (pointer) info to the DBHdr structure
    CopyMemory DBHdr, ByVal lParamIn, LenB(DBHdr)
    Info$ = "----------------------------------------------------------------------" & vbCrLf
    Info$ = Info$ & "=======DEV_BROADCAST_HDR========" & vbCrLf
    Info$ = Info$ & "dbch_devicetype=" & DBHdr.dbch_devicetype & " (" & Hex$(DBHdr.dbch_devicetype) & "H) : "
    Info$ = Info$ & "dbch_reserved=" & DBHdr.dbch_reserved & " (" & Hex$(DBHdr.dbch_reserved) & "H) : "
    Info$ = Info$ & "dbch_size=" & DBHdr.dbch_size & " (" & Hex$(DBHdr.dbch_size) & "H)" & vbCrLf
    Select Case DBHdr.dbch_devicetype
        Case Is = DBT_DEVTYP_VOLUME
            'Logical volume. This structure is a DEV_BROADCAST_VOLUME structure.
            'copy the lparam (pointer) info to the DBVol structure
            CopyMemory DBVol, ByVal lParamIn, LenB(DBVol)
            Info$ = Info$ & "=======DEV_BROADCAST_VOLUME========" & vbCrLf
            Info$ = Info$ & "dbcv_devicetype=" & DBVol.dbcv_devicetype & " (" & Hex$(DBVol.dbcv_devicetype) & "H)" & vbCrLf
            Info$ = Info$ & "dbcv_flags=" & DBVol.dbcv_flags & " (" & Hex$(DBVol.dbcv_flags) & "H)" & vbCrLf
            Info$ = Info$ & "dbcv_reserved=" & DBVol.dbcv_reserved & " (" & Hex$(DBVol.dbcv_reserved) & "H)" & vbCrLf
            Info$ = Info$ & "dbcv_size=" & DBVol.dbcv_size & " (" & Hex$(DBVol.dbcv_size) & "H)" & vbCrLf
            Info$ = Info$ & "dbcv_unitmask=" & DBVol.dbcv_unitmask & " (" & Hex$(DBVol.dbcv_unitmask) & "H)" & vbCrLf
            'from the DeviceNotification example
'            ' Read end of DEV_BROADCAST_VOLUME structure
'            Call GetDWord(ByVal (lParam + Len(DevBroadcastHeader)), UnitMask)
'            Call GetWord(ByVal (lParam + Len(DevBroadcastHeader) + 4), Flags)
'            DriveLetters = UnitMaskToString(UnitMask)
'            Info$ = ChangeType$
'            For LoopDrives = 1 To Len(DriveLetters) ' Print a message for each drive
'                Info$ = Info$ & " Drive " & Mid$(DriveLetters, LoopDrives, 1) & " " & IIf(wParam = DBT_DEVICEARRIVAL, "Inserted", "Ejected") & " (" & DriveTypeToString(GetDriveType(Mid$(DriveLetters, LoopDrives, 1) & ":\")) & ")" & vbCrLf
'            Next LoopDrives
        Case Is = DBT_DEVTYP_OEM '0
            'OEM- or IHV-defined device type.
            'This structure is a DEV_BROADCAST_OEM structure.
            Info$ = Info$ & "=======DBT_DEVTYP_OEM========" & vbCrLf
            Info$ = Info$ & "This structure is a DEV_BROADCAST_OEM structure."
        Case Is = DBT_DEVTYP_PORT ' &H3&
            'Port device (serial or parallel).
            'This structure is a DEV_BROADCAST_PORT structure.
            CopyMemory DBPort, ByVal lParamIn, LenB(DBPort)
            Info$ = Info$ & "=======DEV_BROADCAST_PORT========" & vbCrLf
            Info$ = Info$ & "dbcp_devicetype=" & DBPort.dbcp_devicetype & " (" & Hex$(DBPort.dbcp_devicetype) & "H)" & vbCrLf
            Info$ = Info$ & "dbcp_name=" & DBPort.dbcp_name & " (" & Hex$(DBPort.dbcp_name) & "H)" & vbCrLf
            PortName = Space(255)
            CopyMemory ByVal PortName, DBPort.dbcp_name, 255
            PortName = TrimNull(PortName)
            Info$ = Info$ & "Friendly Name of Port is " & PortName & vbCrLf
            Info$ = Info$ & "dbcp_reserved=" & DBPort.dbcp_reserved & " (" & Hex$(DBPort.dbcp_reserved) & "H)" & vbCrLf
            Info$ = Info$ & "dbcp_size=" & DBPort.dbcp_size & " (" & Hex$(DBPort.dbcp_size) & "H)" & vbCrLf
        Case Is = DBT_DEVTYP_DEVICEINTERFACE ' &H5&
            'Class of devices.This structure is a DEV_BROADCAST_DEVICEINTERFACE structure.
            CopyMemory DBInter, ByVal lParamIn, LenB(DBInter)
'            ' Read end of DEV_BROADCAST_DEVICEINTERFACE structure
'            Call CopyMemoryDBDevInterface(DBInter, ByVal (lParam + Len(DevBroadcastHeader)), Len(DBInter))
'            Call RtlMoveMemory(DeviceGUID, ByVal (lParam + Len(DevBroadcastHeader)), Len(DeviceGUID))
'            Call GetDWord(ByVal (lParam + Len(DevBroadcastHeader) + Len(DeviceGUID)), DeviceNamePtr)
'            Info$ = ChangeType$
            Info$ = Info$ & "=======DEV_BROADCAST_DEVICEINTERFACE========" & vbCrLf
            Info$ = Info$ & " Device GUID: " & GUIDToString(DeviceGUID) & ", name: """ & StringFromPointer(DeviceNamePtr, 1024) & """" & vbCrLf
            Info$ = Info$ & "DBInter.dbcc_classguid:" & GUIDToString(DBInter.dbcc_classguid) & vbCrLf
            Info$ = Info$ & "   DBInter.dbcc_devicetype:" & DBInter.dbcc_devicetype & vbCrLf
            Info$ = Info$ & "   DBInter.dbcc_reserved:" & DBInter.dbcc_reserved & vbCrLf
            Info$ = Info$ & "   DBInter.dbcc_size:" & DBInter.dbcc_size & vbCrLf
            Info$ = Info$ & "   DBInter.dbcc_name - Pointer =" & DBInter.dbcc_name & vbCrLf
            Info$ = Info$ & "   DBInter.dbcc_name: " & StringFromPointer(DBInter.dbcc_name, 1024) & vbCrLf
PortName = Space(255)
            CopyMemory ByVal PortName, DBInter.dbcc_name, 255
            PortName = TrimNull(PortName)
            PortName = CopyStringA(DBInter.dbcc_name)
            Info$ = Info$ & "Friendly Name of Port is " & PortName & vbCrLf
                    
        
        Case Is = DBT_DEVTYP_HANDLE ' &H6&
            'File system handle. This structure is a DEV_BROADCAST_HANDLE structure.
            'If dbch_devicetype is DBT_DEVTYP_HANDLE, the event data is really a pointer
            'to a DEV_BROADCAST_HANDLE structure.
            CopyMemory DBHandle, ByVal lParamIn, LenB(DBHandle)
            Info$ = Info$ & "=======DEV_BROADCAST_HANDLE========" & vbCrLf
            Info$ = Info$ & "dbch_devicetype=" & DBHandle.dbch_devicetype & " (" & Hex$(DBHandle.dbch_devicetype) & "H)" & vbCrLf
            Info$ = Info$ & "dbch_nameoffset=" & DBHandle.dbch_nameoffset & vbCrLf
            Info$ = Info$ & "dbch_handle=" & DBHandle.dbch_handle & vbCrLf
            Info$ = Info$ & "dbch_hdevnotify=" & DBHandle.dbch_hdevnotify & vbCrLf
            Info$ = Info$ & "dbch_nameoffset=" & DBHandle.dbch_nameoffset & vbCrLf
            Info$ = Info$ & "dbch_eventguid.=" & GUIDToString(DeviceGUID)
'            Info$ = Info$ & "dbch_eventguid.ab=" & DBHandle.dbch_eventguid.ab & vbCrLf
'            Info$ = Info$ & "dbch_eventguid.ac=" & DBHandle.dbch_eventguid.ac & vbCrLf
'            Info$ = Info$ & "dbch_eventguid.ad=" & DBHandle.dbch_eventguid.ad & vbCrLf
'            For i = 0 To 7
'                Info$ = Info$ & "dbch_eventguid.ae(" & i & ")=" & DBHandle.dbch_eventguid.ae(i) & vbCrLf
'            Next i
            Info$ = Info$ & "dbch_reserved=" & DBHandle.dbch_reserved & " (" & Hex$(DBHandle.dbch_reserved) & "H)" & vbCrLf
            Info$ = Info$ & "dbch_size=" & DBHandle.dbch_size & " (" & Hex$(DBHandle.dbch_size) & "H)" & vbCrLf
        Case Else
            Info$ = Info$ & "Unknown DBHdr.dbch_devicetype = " & DBHdr.dbch_devicetype
    End Select
    Select Case wParamIn
        Case DBT_DEVICEARRIVAL
            'See if a CD-ROM or DVD was inserted into a drive.
            If DBHdr.dbch_devicetype = DBT_DEVTYP_VOLUME Then
                If (DBVol.dbcv_flags And DBTF_MEDIA) = DBTF_MEDIA Then
                    Info$ = Info$ & "New media inserted in drive " & GetVolumeLetter(DBVol.dbcv_unitmask)
                ElseIf (DBVol.dbcv_flags And DBTF_NET) = DBTF_NET Then
                    Info$ = Info$ & "New network Drive " & GetVolumeLetter(DBVol.dbcv_unitmask) & " added"
                Else
                    Info$ = Info$ & "Unknown DBVol.dbcv_flags = " & DBVol.dbcv_flags
                    Info$ = Info$ & "(" & Hex$(DBVol.dbcv_flags) & "H) from drive "
                    Info$ = Info$ & GetVolumeLetter(DBVol.dbcv_unitmask)
                End If
            End If
        Case DBT_DEVICEREMOVECOMPLETE
            'See if a CD-ROM or DVD was removed from a drive.
            If DBHdr.dbch_devicetype = DBT_DEVTYP_VOLUME Then
                If (DBVol.dbcv_flags And DBTF_MEDIA) = DBTF_MEDIA Then
                    Info$ = Info$ & "Media removed from drive " & GetVolumeLetter(DBVol.dbcv_unitmask)
                ElseIf (DBVol.dbcv_flags And DBTF_NET) = DBTF_NET Then
                    Info$ = Info$ & "New network Drive " & GetVolumeLetter(DBVol.dbcv_unitmask) & " added"
                Else
                    Info$ = Info$ & "Unknown DBVol.dbcv_flags = " & DBVol.dbcv_flags
                    Info$ = Info$ & "(" & Hex$(DBVol.dbcv_flags) & "H) from drive "
                    Info$ = Info$ & GetVolumeLetter(DBVol.dbcv_unitmask)
                End If
            Else
                Info$ = Info$ & "Unknown DBHdr.dbch_devicetype = " & DBHdr.dbch_devicetype
            End If
        Case Else
            Info$ = Info$ & "Unknown Param for device change" & wParamIn
    End Select
    If Len(Info$) > 0 Then
        Debug.Print Info$
        Form1.Text1.SelText = Info$ & vbCrLf
        'MsgBox Info$, vbExclamation, "ProcessDeviceChange"
    End If
Else
    Debug.Print "lParamIn is Zero - Cannot Process Device Change"
End If
End Sub

Private Function GetVolumeLetter(Mask As Long) As String
'This function takes the mask value passed from DBVol.dbcv_unitmask
'and returns the logical drive letter
'bit 0=drive A (mask=01H); bit 1=B(mask=02H); ... bit 5=E (10H)
'etc. all to way up to bit 26 =Z(2000000H)
Dim DriveCharCode As Integer
DriveCharCode = 65 'start with the code for drive A
Do Until (Mask And 1) = 1
    Mask = Fix(Mask / 2) ' If Fix isn't supported try " Abs(Mask / 2) - 1 "
    DriveCharCode = DriveCharCode + 1
Loop
'Return the ascii character for the code value
GetVolumeLetter = Chr$(DriveCharCode)
End Function

Private Function TrimNull(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
    OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
TrimNull = OriginalStr
End Function

'----------------------------------------------------------------------
'Below is the data from several events
'External Networked E Drive
'=======DEV_BROADCAST_HDR========
'dbch_devicetype=2 (2H) : dbch_reserved=0 (0H) : dbch_size=20 (14H)
'=======DEV_BROADCAST_VOLUME========
'dbcv_devicetype=2 (2H)
'dbcv_flags=64684034 (3DB0002H)
'dbcv_reserved=0 (0H)
'dbcv_size=20 (14H)
'dbcv_unitmask=33554432 (2000000H)
'Unknown DBVol.dbcv_flags = 64684034(3DB0002H) from drive Z
'DBT_DeviceRemoveComplete - lParam = 1310180(13FDE4H)
'----------------------------------------------------------------------
'=======DEV_BROADCAST_HDR========
'dbch_devicetype=2 (2H) : dbch_reserved=0 (0H) : dbch_size=20 (14H)
'=======DEV_BROADCAST_VOLUME========
'dbcv_devicetype=2 (2H)
'dbcv_flags=-2141978622 (80540002H)
'dbcv_reserved=0 (0H)
'dbcv_size=20 (14H)
'dbcv_unitmask=33554432 (2000000H)
'Unknown DBVol.dbcv_flags = -2141978622(80540002H) from drive Z
'DBT_DeviceRemoveComplete - lParam = 1310180(13FDE4H)
'----------------------------------------------------------------------
'=======DEV_BROADCAST_HDR========
'dbch_devicetype=2 (2H) dbch_reserved=0 (0H) dbch_size=20 (14H)
'=======DEV_BROADCAST_VOLUME========
'dbcv_devicetype=2 (2H)
'dbcv_flags=-2142306302 (804F0002H)
'dbcv_reserved=0 (0H)
'dbcv_size=20 (14H)
'dbcv_unitmask=16 (10H)
'Unknown DBVol.dbcv_flags = -2142306302 from drive E
'DBT_DeviceArrival - lParam = 1310180(13FDE4H)
'----------------------------------------------------------------------
'networked C drive
'=======DEV_BROADCAST_HDR========
'dbch_devicetype=2 (2H) dbch_reserved=0 (0H) dbch_size=20 (14H)
'=======DEV_BROADCAST_VOLUME========
'dbcv_devicetype=2 (2H)
'dbcv_flags=-2142306302 (804F0002H)
'dbcv_reserved=0 (0H)
'dbcv_size=20 (14H)
'dbcv_unitmask=33554432 (2000000H)
'Unknown DBVol.dbcv_flags = -2142306302 from drive Z
'DBT_DeviceArrival - lParam = 1310180(13FDE4H)
'Unknown Event - wParam = 7(7H)lParam = 0(0H)
'----------------------------------------------------------------------
'Thumb Drive
'=======DEV_BROADCAST_HDR========
'dbch_devicetype=2 (2H) dbch_reserved=0 (0H) dbch_size=20 (14H)
'=======DEV_BROADCAST_VOLUME========
'dbcv_devicetype=2 (2H)
'dbcv_flags=404226048 (18180000H)
'dbcv_reserved=0 (0H)
'dbcv_size=20 (14H)
'dbcv_unitmask=32 (20H)
'Unknown DBVol.dbcv_flags = 404226048 from drive F
'DBT_DeviceArrival - lParam = 1310180(13FDE4H)
'----------------------------------------------------------------------
'Removed Thumb
'=======DEV_BROADCAST_HDR========
'dbch_devicetype=2 (2H) dbch_reserved=0 (0H) dbch_size=20 (14H)
'=======DEV_BROADCAST_VOLUME========
'dbcv_devicetype=2 (2H)
'dbcv_flags=0 (0H)
'dbcv_reserved=0 (0H)
'dbcv_size=20 (14H)
'dbcv_unitmask=32 (20H)
'Unknown DBVol.dbcv_flags = 0 from drive F
'DBT_DeviceRemoveComplete - lParam = 1310180(13FDE4H)
'
'----------------------------------------------------------------------
'keyspan usb serial port
'=======DEV_BROADCAST_HDR========
'dbch_devicetype=3 (3H) : dbch_reserved=0 (0H) : dbch_size=22 (16H)
'=======DEV_BROADCAST_PORT========
'dbcp_devicetype=3 (3H)
'dbcp_name=843927363 (324D4F43H)
'Friendly Name of Port is COM2
'dbcp_reserved=0 (0H)
'dbcp_size=22 (16H)
'DBT_DeviceArrival - lParam = 1898880(1CF980H)

Private Function UnitMaskToString(ByVal inUnitMask As Long) As String
Dim LoopBits As Long
For LoopBits = 0 To 30
    If (inUnitMask And (2 ^ LoopBits)) Then UnitMaskToString = UnitMaskToString & Chr$(Asc("A") + LoopBits)
Next LoopBits
End Function

Private Function GUIDToString(ByRef inGUID As Guid) As String
Dim RetBuf As String
Dim GUILen As Long
Const BufLen As Long = 80
RetBuf = Space$(BufLen)
GUILen = StringFromGUID2(inGUID, RetBuf, BufLen)
If (GUILen) Then
    GUIDToString = StrConv(Left$(RetBuf, (GUILen - 1) * 2), vbFromUnicode)
End If
End Function

Public Function CopyStringA(ByVal inPtr As Long) As String
Dim BufLen As Long
BufLen = lstrlenA(inPtr)
'Debug.Print "BufLen="; BufLen
If (BufLen > 0) Then
    CopyStringA = Space$(BufLen)
    Call lstrcpyA(CopyStringA, inPtr)
End If
End Function

Public Function StringFromPointer(lpString As Long, lMaxLength As Long) As String
Dim sRet As String
Dim lret As Long
If lpString = 0 Then
    StringFromPointer = ""
    Exit Function
End If
lret = lstrlen(lpString)
'Debug.Print "lstrlen="; lret
If lret < lMaxLength Then
    lMaxLength = lret
End If
If IsBadStringPtrByLong(lpString, lMaxLength) Then
    ' An error has occured - do not attempt to use this pointer
    MsgBox "Attempt to read bad string pointer: " & lpString, , "StringFromPointer Error:" & Err.LastDllError
    StringFromPointer = ""
    Exit Function
End If
' Pre-initialise the return string...
sRet = Space$(lMaxLength)
Call lstrcpyToBuffer(sRet, lpString)
'CopyMemory ByVal sRet, ByVal lpString, ByVal Len(sRet)
If Err.LastDllError = 0 Then
    If InStr(sRet, Chr$(0)) > 0 Then
        sRet = Left$(sRet, InStr(sRet, Chr$(0)) - 1)
    End If
End If
StringFromPointer = sRet
'Debug.Print "success - "; sRet
End Function

Private Function DriveTypeToString(ByVal inDriveType As Long) As String
Select Case inDriveType
    Case DRIVE_NO_ROOT_DIR
        DriveTypeToString = "No root directory" '??
    Case DRIVE_REMOVABLE
        DriveTypeToString = "Removable"
    Case DRIVE_FIXED
        DriveTypeToString = "Fixed"
    Case DRIVE_REMOTE
        DriveTypeToString = "Remote"
    Case DRIVE_CDROM
        DriveTypeToString = "CD-ROM"
    Case DRIVE_RAMDISK
        DriveTypeToString = "RAM disk"
    Case Else
        DriveTypeToString = "[ Unknown ]"
End Select
End Function

