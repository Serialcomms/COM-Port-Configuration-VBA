Attribute VB_Name = "COM_PORT_ENUM_SETTINGS"
'
' https://github.com/Serialcomms/COM-Port-Configuration-VBA
'
' https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/
' https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-commconfigdialoga
'
Option Explicit

Private Type DEVICE_CONTROL_BLOCK

             LENGTH_DCB         As Long
             BAUD_RATE          As Long
             BIT_FIELD          As Long
             RESERVED_0         As Integer
             LIMIT_XON          As Integer
             LIMIT_XOFF         As Integer
             BYTE_SIZE          As Byte
             PARITY             As Byte
             STOP_BITS          As Byte
             CHAR_XON           As Byte
             CHAR_XOFF          As Byte
             CHAR_ERROR         As Byte
             CHAR_EOF           As Byte
             CHAR_EVENT         As Byte
             RESERVED_1         As Integer
End Type

Private Type COMM_CONFIG
    
             Size                As Long
             Version             As Integer
             Reserved            As Integer
             DCB                 As DEVICE_CONTROL_BLOCK
             Provider_SubType    As Long
             Provider_Offset     As Long
             Provider_Size       As Long
             Provider_Data       As String * 1
End Type

Private Type COM_PORT_PROFILE

             Name                As String
             Label               As String
             Settings            As String
             Handle              As LongPtr
             New_Comm_Config     As COMM_CONFIG
             Port_Comm_Config    As COMM_CONFIG
             New_DCB             As DEVICE_CONTROL_BLOCK
             Port_DCB            As DEVICE_CONTROL_BLOCK
End Type

Private Com_Port_Count As Long
Private Com_Port_Index As Long
Private Com_Port_Ribbon As IRibbonUI

Private MESSAGE_BOX_TEXT As String
Private MESSAGE_BOX_TITLE As String
Private MESSAGE_BOX_RESULT As Long
Private MESSAGE_BOX_BUTTONS As Long

Private Const TEXT_COM As String = "COM"
Private Const TEXT_DASH As String = "-"
Private Const TEXT_NEWLINES As String = vbCrLf & vbCrLf
Private Const TEXT_COM_PORT As String = "  COM Port "
Private Const TEXT_NO_COM_PORT As String = "  No COM Port "
Private Const TEXT_NO_COM_PORTS As String = "  No COM Ports "

Private Declare PtrSafe Function Get_Com_Ports Lib "KernelBase.dll" Alias "GetCommPorts" _
(ByRef Port_Array As Long, ByVal Array_Length As Long, ByRef Port_Count As Long) As Long

Private Declare PtrSafe Function Get_Comm_Config Lib "kernel32" Alias "GetCommConfig" _
(ByVal Port_Handle As LongPtr, ByRef CC As COMM_CONFIG, ByVal CC_Length As Long) As Long

Private Declare PtrSafe Function Get_Comm_Default Lib "kernel32" Alias "GetDefaultCommConfigA" _
(ByVal Port_Name As String, ByRef CC As COMM_CONFIG, ByVal CC_Length As Long) As Long

Private Declare PtrSafe Function Port_Config_Dialogue Lib "kernel32" Alias "CommConfigDialogA" _
(ByVal Port_Name As String, ByVal Window_Handle As LongPtr, ByRef CC As COMM_CONFIG) As Long

Private Declare PtrSafe Function Query_Port_DCB Lib "Kernel32.dll" Alias "GetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Apply_Port_DCB Lib "Kernel32.dll" Alias "SetCommState" (ByVal Port_Handle As LongPtr, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean
Private Declare PtrSafe Function Build_Port_DCB Lib "Kernel32.dll" Alias "BuildCommDCBA" (ByVal Config_Text As String, ByRef Port_DCB As DEVICE_CONTROL_BLOCK) As Boolean

Private Declare PtrSafe Function Com_Port_Create Lib "Kernel32.dll" Alias "CreateFileA" _
(ByVal Port_Name As String, ByVal PORT_ACCESS As Long, ByVal SHARE_MODE As Long, ByVal SECURITY_ATTRIBUTES_NULL As Any, _
 ByVal CREATE_DISPOSITION As Long, ByVal FLAGS_AND_ATTRIBUTES As Long, Optional TEMPLATE_FILE_HANDLE_NULL) As LongPtr
 
Private Declare PtrSafe Function Com_Port_Close Lib "Kernel32.dll" Alias "CloseHandle" (ByVal Port_Handle As LongPtr) As Boolean

Private Declare PtrSafe Sub Kernel_Sleep_Milliseconds Lib "Kernel32.dll" Alias "Sleep" (ByVal Sleep_Milliseconds As Long)

Private Temp_Text As String

Private Last_Port_Scan As String

Private COM_PORTS() As COM_PORT_PROFILE

Private Const LONG_0 As Long = 0
Private Const LONG_1 As Long = 1
Private Const LONG_2 As Long = 2
Private Const LONG_3 As Long = 3
Private Const LONG_4 As Long = 4
Private Const HANDLE_INVALID As LongPtr = -1
Private Const Max_Port_Count As Long = 255
Private Temp_Port_Numbers(LONG_1 To Max_Port_Count) As Long
'

Private Sub SET_PORT_DIALOGUE(Port_Index As Long)

Dim DLL_Error As Long

With COM_PORTS(Port_Index)

 MESSAGE_BOX_TITLE = .Label

 .New_Comm_Config.DCB = .New_DCB
 .New_Comm_Config.Size = LenB(.New_Comm_Config)

 Port_Config_Dialogue .Name, LONG_0, .New_Comm_Config
 
 DLL_Error = Err.LastDllError

End With

Select Case DLL_Error
        
 Case LONG_0:   APPLY_PORT_SETTINGS Port_Index
                         
 Case 87:       MESSAGE_BOX_TEXT = "COM PORT ERROR": MESSAGE_BOX_BUTTONS = vbCritical + vbOKOnly
              ' DLL Error 87 may be returned by COM port types which do not support settings changes - e.g. some virtual COM ports
  
 Case 1223:     MESSAGE_BOX_TEXT = "Change Setttings Cancelled": MESSAGE_BOX_BUTTONS = vbInformation + vbOKOnly
    
 Case Else:     MESSAGE_BOX_TEXT = "Unknown Error": MESSAGE_BOX_BUTTONS = vbCritical + vbOKOnly
          
End Select

If DLL_Error <> LONG_0 Then MsgBox MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE
    
End Sub

Private Sub APPLY_PORT_SETTINGS(Port_Index As Long)

Dim DLL_Error As Long
Dim Old_Settings As String
Dim New_Settings As String
Dim Apply_Result As Boolean

With COM_PORTS(Port_Index)

MESSAGE_BOX_TITLE = .Label

Old_Settings = IIf(Len(.Settings) = LONG_0, "Unknown", .Settings)

With .New_Comm_Config.DCB
        
 New_Settings = SETTINGS_TO_STRING(.BAUD_RATE, .BYTE_SIZE, .PARITY, .STOP_BITS)
    
End With

MESSAGE_BOX_TEXT = "Apply settings " & New_Settings & " to port " & .Name & " ? "
MESSAGE_BOX_BUTTONS = vbQuestion + vbOKCancel
MESSAGE_BOX_RESULT = MsgBox(MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE)

Select Case MESSAGE_BOX_RESULT

Case vbOK

If OPEN_COM_PORT(Port_Index) Then

    Apply_Result = Apply_Port_DCB(.Handle, .New_Comm_Config.DCB)

    If Apply_Result Then
          
        Temp_Text = "Updated" & TEXT_NEWLINES & "Old Settings = " & Old_Settings
        Temp_Text = Temp_Text & TEXT_NEWLINES & "New Settings = " & New_Settings
   
        MESSAGE_BOX_TEXT = .Name & " " & Temp_Text
        MESSAGE_BOX_BUTTONS = vbInformation + vbOKOnly
        MsgBox MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE
        
        .Settings = New_Settings
        
        Com_Port_Ribbon.InvalidateControl ("CP_Settings")
     
     Else
     
        DLL_Error = Err.LastDllError
        
        Temp_Text = .Name & TEXT_NEWLINES & "Old Settings = " & Old_Settings & vbCrLf
        
        MESSAGE_BOX_TEXT = "Error applying settings to " & Temp_Text & vbCrLf & "DLL Error = " & DLL_Error
        MESSAGE_BOX_BUTTONS = vbCritical + vbOKOnly
        MsgBox MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE
     
     End If
   
     Com_Port_Close .Handle
     
     .Handle = LONG_0
     
Else ' com port open fail - e.g. DLL error 5, port in use.

End If

Case vbCancel

        MESSAGE_BOX_TEXT = "Change Settings Cancelled"
        MESSAGE_BOX_BUTTONS = vbInformation + vbOKOnly
        MsgBox MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE

Case Else

End Select

End With

End Sub

Public Function SETTINGS_TO_STRING(DCB_Baud As Long, DCB_Byte As Byte, DCB_Parity As Byte, DCB_Stop As Byte) As String

Dim Settings_String As String

Settings_String = Settings_String & DCB_Baud & TEXT_DASH
Settings_String = Settings_String & DCB_Byte & TEXT_DASH
Settings_String = Settings_String & CONVERT_PARITY(DCB_Parity) & TEXT_DASH
Settings_String = Settings_String & CONVERT_STOPBITS(DCB_Stop)

SETTINGS_TO_STRING = Settings_String

End Function

Public Function SHOW_PORT_DEFAULT(Optional Port_Name As String)

Dim Config_Result As Boolean
Dim Default_Settings As String
Dim Port_Default_Config As COMM_CONFIG
Dim Port_Default_DCB As DEVICE_CONTROL_BLOCK

Const TEXT_PORT_ERROR As String = "PORT-ERROR"

Port_Default_Config.Size = LenB(Port_Default_Config)
Port_Default_Config.DCB = Port_Default_DCB

If Len(Port_Name) > LONG_3 Then

    Config_Result = Get_Comm_Default(Port_Name, Port_Default_Config, Port_Default_Config.Size)

    If Config_Result Then
    
     With Port_Default_Config.DCB
        
        Default_Settings = SETTINGS_TO_STRING(.BAUD_RATE, .BYTE_SIZE, .PARITY, .STOP_BITS)
        
     End With
    
    End If
    
Else

    Default_Settings = TEXT_PORT_ERROR

End If

SHOW_PORT_DEFAULT = Default_Settings

End Function

Private Function OPEN_COM_PORT(Port_Index As Long) As Boolean

Dim Device_Path As String
Dim Open_Handle As LongPtr
Dim Open_Result As Boolean

Const OPEN_EXISTING As Long = LONG_3
Const OPEN_EXCLUSIVE As Long = LONG_0
Const SYNCHRONOUS_MODE As Long = LONG_0

Const GENERIC_RW As Long = &HC0000000
Const DEVICE_PREFIX As String = "\\.\"
        
Device_Path = DEVICE_PREFIX & COM_PORTS(Port_Index).Name

Open_Handle = Com_Port_Create(Device_Path, GENERIC_RW, OPEN_EXCLUSIVE, LONG_0, OPEN_EXISTING, SYNCHRONOUS_MODE)

Open_Result = Not (Open_Handle = HANDLE_INVALID)

If Not Open_Result Then

    COM_PORTS(Port_Index).Handle = HANDLE_INVALID
    
    Temp_Text = COM_PORTS(Port_Index).Name & TEXT_NEWLINES & "DLL Error = " & Err.LastDllError

    MESSAGE_BOX_TEXT = "Error Opening Port " & Temp_Text
    MESSAGE_BOX_TITLE = COM_PORTS(Port_Index).Label & " Error"
    MESSAGE_BOX_BUTTONS = vbCritical + vbOKOnly
    MsgBox MESSAGE_BOX_TEXT, MESSAGE_BOX_BUTTONS, MESSAGE_BOX_TITLE

Else
 
    COM_PORTS(Port_Index).Handle = Open_Handle
    
    Query_Port_DCB Open_Handle, COM_PORTS(Port_Index).Port_DCB
    
    With COM_PORTS(Port_Index).Port_DCB
    
     COM_PORTS(Port_Index).Settings = SETTINGS_TO_STRING(.BAUD_RATE, .BYTE_SIZE, .PARITY, .STOP_BITS)
    
    End With
    
End If

OPEN_COM_PORT = Open_Result

End Function

Private Function Query_Com_Ports() As Long

Dim Port_Number As Long
Dim Port_Ordinal As Long

Last_Port_Scan = Date$ & " " & Time$

Get_Com_Ports Temp_Port_Numbers(LONG_1), Max_Port_Count, Com_Port_Count

ReDim COM_PORTS(LONG_0 To Com_Port_Count)

COM_PORTS(LONG_0).Label = IIf(Com_Port_Count = LONG_0, TEXT_NO_COM_PORTS, TEXT_NO_COM_PORT)

If Com_Port_Count > LONG_0 Then

    For Port_Ordinal = LONG_1 To Com_Port_Count
    
        Port_Number = Temp_Port_Numbers(Port_Ordinal)
    
        COM_PORTS(Port_Ordinal).Name = TEXT_COM & Port_Number
        COM_PORTS(Port_Ordinal).Label = TEXT_COM_PORT & Port_Number
        COM_PORTS(Port_Ordinal).Settings = vbNullString

    Next Port_Ordinal
    
End If

Query_Com_Ports = Com_Port_Count

End Function

Public Function Read_Ribbon_Combo() As String

Application.Volatile

If Com_Port_Index = LONG_0 Then

    Read_Ribbon_Combo = vbNullString

Else

    Read_Ribbon_Combo = COM_PORTS(Com_Port_Index).Name
   
End If

End Function
                
Private Sub InitPortRibbon(Port_Ribbon As IRibbonUI)                                       ' Ribbon Callback for customUI.onLoad

Set Com_Port_Ribbon = Port_Ribbon

Query_Com_Ports

End Sub

Private Sub PortScan(Button_Control As IRibbonControl)                                     ' Ribbon Callback for CP_Button onAction

Query_Com_Ports

Com_Port_Ribbon.InvalidateControl ("CP_Button")

Application.Calculate

End Sub

Private Sub GetButtonLabel(Control As IRibbonControl, ByRef ButtonLabel)                   ' Ribbon Callback for CP_Button getLabel

Const TEXT_SELECT As String = " Select COM Port "
Const TEXT_DETECT As String = " Detect COM Ports"

ButtonLabel = IIf(Com_Port_Count = LONG_0, TEXT_DETECT, TEXT_SELECT)

Com_Port_Ribbon.InvalidateControl ("CP_Selector")

End Sub

Private Sub GetSupertipText(Control As IRibbonControl, ByRef SupertipText)                 ' Ribbon Callback for CP_Button getSupertipText

Dim Scan_Text As String

Const TEXT_PORTS_AVAILABLE As String = vbCrLf & "Com Ports Available = "

Const TEXT_NO_PORTS_FOUND As String = vbCrLf & "No Com ports available " & TEXT_NEWLINES & "Click to rescan for new Com ports"

Scan_Text = TEXT_NEWLINES & "Last Port Scan " & IIf(Len(Last_Port_Scan) = LONG_0, "Unknown", Last_Port_Scan)

SupertipText = IIf(Com_Port_Count = LONG_0, TEXT_NO_PORTS_FOUND, TEXT_PORTS_AVAILABLE & Com_Port_Count) & Scan_Text

End Sub

Private Sub GetPortCount(Control As IRibbonControl, ByRef DropDown_Entries)                 ' Ribbon Callback for CP_Selector getPortCount

DropDown_Entries = LONG_1 + Query_Com_Ports

Com_Port_Ribbon.InvalidateControl ("CP_Settings")

End Sub

Private Sub AddPortID(Control As IRibbonControl, Index As Integer, ByRef PortID)            ' Ribbon Callback for CP_Selector getPortID

PortID = "Port_ID_" & Index

End Sub

Private Sub AddPortLabel(Control As IRibbonControl, Index As Integer, ByRef PortLabel)      ' Ribbon Callback for CP_Selector getPortLabel

PortLabel = COM_PORTS(Index).Label

End Sub

Private Sub GetPortIndex(Control As IRibbonControl, id As String, PortIndex As Long)        ' Ribbon Callback for CP_Selector onChange

Com_Port_Index = PortIndex

Com_Port_Ribbon.InvalidateControl ("CP_Settings")

Application.Calculate

End Sub

Private Sub EnableSettings(Control As IRibbonControl, ByRef returnedVal)                    ' Ribbon Callback for CP_Settings getEnabled

returnedVal = IIf(Com_Port_Index = LONG_0, False, True)

End Sub

Private Sub PortSettings(Control As IRibbonControl)                                         ' Ribbon Callback for CP_Settings onAction

SET_PORT_DIALOGUE Com_Port_Index

End Sub

Private Sub GetSettingsText(Control As IRibbonControl, ByRef returnedVal)                   ' Ribbon Callback for CP_Settings getSupertip

returnedVal = vbCrLf & "Run Windows Com Port Settings"

End Sub

Private Sub GetSettingsLabel(Control As IRibbonControl, ByRef returnedVal)                  ' Ribbon Callback for CP_Settings getLabel

Dim Label_Suffix As String

If Com_Port_Count = LONG_0 Then

    returnedVal = vbNullString

Else

    With COM_PORTS(Com_Port_Index)

     Label_Suffix = IIf(Len(.Settings) = LONG_0, "Settings", .Settings)
    
     returnedVal = .Name & vbCrLf & Label_Suffix
    
    End With

End If

End Sub

Public Function CONVERT_PARITY(DCB_Parity As Byte) As String

Dim Parity_Text As String

Select Case DCB_Parity

Case LONG_0:    Parity_Text = "N"
Case LONG_1:    Parity_Text = "O"
Case LONG_2:    Parity_Text = "E"
Case LONG_3:    Parity_Text = "M"
Case LONG_4:    Parity_Text = "S"

Case Else:      Parity_Text = "?"

End Select

CONVERT_PARITY = Parity_Text

End Function

Public Function CONVERT_STOPBITS(DCB_STOPBITS As Byte) As String

Dim Stop_Text As String

Select Case DCB_STOPBITS

Case LONG_0:    Stop_Text = "1"
Case LONG_2:    Stop_Text = "2"
Case LONG_1:    Stop_Text = "1.5"

Case Else:      Stop_Text = "?"

End Select

CONVERT_STOPBITS = Stop_Text

End Function
