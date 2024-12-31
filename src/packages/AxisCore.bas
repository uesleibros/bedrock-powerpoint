Attribute VB_Name = "AxisCore"
' **************************************************************************
' AxisCore - Joystick Axis Direction Detection
' **************************************************************************
' This module handles the detection of joystick analog stick directions.
' It processes the raw axis values for the X and Y axes and returns
' the appropriate joystick axis direction.

' The module ensures accurate detection of all directions (up, down, left, right)
' and diagonal directions (up-right, up-left, down-right, down-left),
' based on the raw input values from the joystick axes.
'
' Designed for integration with joystick handling in VBA projects.
' It provides an efficient mechanism for converting raw axis values into
' meaningful directional inputs for game development, simulation, or other
' control systems that require precise joystick movement tracking.
'
' --------------------------------------------------------------------------
' Team: Bedrock
' Last Update: 12/30/2024
' --------------------------------------------------------------------------
' References:
' - https://learn.microsoft.com/en-us/windows/win32/api/joystickapi/nf-joystickapi-joygetposex
' - https://stackoverflow.com/questions/23739325/why-joygetpos-works-and-joygetposex-does-not
' - https://learn.microsoft.com/en-us/windows/win32/api/joystickapi/
' - https://learn.microsoft.com/en-us/windows/win32/api/joystickapi/ns-joystickapi-joycaps
' - https://learn.microsoft.com/en-us/windows/win32/api/joystickapi/ns-joystickapi-joyinfoex

Option Explicit

Private Declare PtrSafe Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Private Declare PtrSafe Function joyGetNumDevs Lib "winmm.dll" () As Long
Private Declare PtrSafe Function joyGetDevCapsW Lib "winmm.dll" (ByVal uJoyID As Long, ByRef pjc As JOYCAPSW, ByVal cbjc As Long) As Long

Private Const MMSYSERR_NOERROR = 0
Private Const JOY_RETURNALL = &HFF
Private Const MAX_JOYSTICKID = 20
Private Const DEFAULT_DEADZONE = 2000
Private Const DEFAULT_AXIS_CENTER = 32767
Private Const POV_CENTERED = -1

Private JoystickStates(MAX_JOYSTICKID) As JOYINFOEX
Private LastJoystickStates(MAX_JOYSTICKID) As JOYINFOEX
Private ConnectedJoysticks(MAX_JOYSTICKID) As Boolean
Private LastDPadState As DPadDirection

Private Type JOYCAPSW
    ManufacturerID As Integer
    ProductID As Integer
    ProductName As String * 32
    XMin As Long
    XMax As Long
    YMin As Long
    YMax As Long
    ZMin As Long
    ZMax As Long
    ButtonCount As Long
    PeriodMin As Long
    PeriodMax As Long
    RMin As Long
    RMax As Long
    UMin As Long
    UMax As Long
    VMin As Long
    VMax As Long
    Caps As Long
    MaxAxes As Long
    AxisCount As Long
    MaxButtons As Long
    RegistryKey As String * 32
    OEMDriver As String * 260
End Type

Private Type JOYINFOEX
    Size As Long
    Flags As Long
    XPosition As Long
    YPosition As Long
    ZPosition As Long
    RPosition As Long
    UPosition As Long
    VPosition As Long
    Buttons As Long
    ButtonPressed As Long
    POV As Long
    Reserved1 As Long
    Reserved2 As Long
End Type

Private Type Joystick
    ID As Long
    Name As String
    Deadzone As Integer
    Connected As Boolean
    State As JOYINFOEX
End Type

Private Type JoystickCollection
    Devices(MAX_JOYSTICKID) As Joystick
    TotalConnected As Byte
End Type

Public Enum JoystickAxis
    LEFT_ANALOG_X = 0
    LEFT_ANALOG_Y = 1
    RIGHT_ANALOG_X = 2
    RIGHT_ANALOG_Y = 3
    TRIGGER_LT = 4
    TRIGGER_RT = 5
End Enum

Public Enum JoystickButton
    BUTTON_X = 0
    BUTTON_A = 1
    BUTTON_B = 2
    BUTTON_Y = 3
    BUTTON_LB = 4
    BUTTON_RB = 5
    BUTTON_LT = 6
    BUTTON_RT = 7
    BUTTON_SELECT = 8
    BUTTON_START = 9
    BUTTON_LS = 10
    BUTTON_RS = 11
End Enum

Public Enum JoystickAxisDirection
    AXIS_DIRECTION_CENTERED = -1
    AXIS_DIRECTION_UP = 0
    AXIS_DIRECTION_RIGHT = 9000
    AXIS_DIRECTION_DOWN = 18000
    AXIS_DIRECTION_LEFT = 27000
    AXIS_DIRECTION_UP_RIGHT = 4500
    AXIS_DIRECTION_UP_LEFT = 31500
    AXIS_DIRECTION_DOWN_RIGHT = 13500
    AXIS_DIRECTION_DOWN_LEFT = 22500
End Enum

Public Enum DPadDirection
    CENTERED = -1
    DIRECTION_UP = 0
    DIRECTION_RIGHT = 9000
    DIRECTION_DOWN = 18000
    DIRECTION_LEFT = 27000
    DIRECTION_UP_RIGHT = 4500
    DIRECTION_UP_LEFT = 31500
    DIRECTION_DOWN_RIGHT = 13500
    DIRECTION_DOWN_LEFT = 22500
End Enum

Public Joysticks As JoystickCollection

Public Sub Initialize()
    InitializeJoysticks
End Sub

Private Sub InitializeJoysticks()
    Dim i As Long, joyInfo As JOYINFOEX, JoyCaps As JOYCAPSW

    Joysticks.TotalConnected = 0

    joyInfo.Size = LenB(joyInfo)
    joyInfo.Flags = JOY_RETURNALL

    For i = 0 To MAX_JOYSTICKID
        If joyGetPosEx(i, joyInfo) = MMSYSERR_NOERROR Then
            InitializeJoystick i, joyInfo
        End If
    Next i
End Sub

Private Sub InitializeJoystick(joyID As Long, joyInfo As JOYINFOEX)
    If Joysticks.Devices(joyID).Connected Then Exit Sub
    Dim NewJoystick As Joystick
    Dim JoyCaps As JOYCAPSW
    Dim result As Long

    result = joyGetDevCapsW(joyID, JoyCaps, LenB(JoyCaps))
    If result = 0 Then
        NewJoystick.ID = joyID
        NewJoystick.Name = Trim(Replace(JoyCaps.ProductName, Chr$(0), ""))
        NewJoystick.Deadzone = DEFAULT_DEADZONE
        NewJoystick.State = joyInfo
        NewJoystick.Connected = True
        Joysticks.Devices(Joysticks.TotalConnected) = NewJoystick
        Joysticks.TotalConnected = Joysticks.TotalConnected + 1
    End If
End Sub

Public Sub UpdateInput()
    Dim i As Long

    For i = 0 To MAX_JOYSTICKID
        LastJoystickStates(i) = JoystickStates(i)
        UpdateJoystickState i
    Next i
End Sub

Private Sub UpdateJoystickState(ByVal joyID As Long)
    Dim joyInfo As JOYINFOEX
    joyInfo.Size = LenB(joyInfo)
    joyInfo.Flags = JOY_RETURNALL

    If joyGetPosEx(joyID, joyInfo) = MMSYSERR_NOERROR Then
        JoystickStates(joyID) = joyInfo
        If Not ConnectedJoysticks(joyID) Then
            InitializeJoystick joyID, joyInfo
            Debug.Print Joysticks.Devices(joyID).Name & "(id:" & Joysticks.Devices(joyID).ID & ") [CONNECTED]"
        End If
        ConnectedJoysticks(joyID) = True
    Else
        If ConnectedJoysticks(joyID) Then
            Debug.Print Joysticks.Devices(joyID).Name & "(id:" & Joysticks.Devices(joyID).ID & ") [DISCONNECTED]"
            Joysticks.Devices(joyID).Connected = False
            Joysticks.TotalConnected = Joysticks.TotalConnected - 1
        End If
        ConnectedJoysticks(joyID) = False
        With JoystickStates(joyID)
            .XPosition = DEFAULT_AXIS_CENTER
            .YPosition = DEFAULT_AXIS_CENTER
            .ZPosition = DEFAULT_AXIS_CENTER
            .RPosition = DEFAULT_AXIS_CENTER
            .UPosition = DEFAULT_AXIS_CENTER
            .VPosition = DEFAULT_AXIS_CENTER
            .Buttons = 0
            .POV = POV_CENTERED
        End With
    End If
End Sub

Public Function IsConnected(ByVal joyID As Long) As Boolean
    IsConnected = ConnectedJoysticks(joyID)
End Function

Public Function IsButtonDown(ByVal joyID As Long, ByVal button As JoystickButton) As Boolean
    If Not ConnectedJoysticks(joyID) Then Exit Function
    
    IsButtonDown = (JoystickStates(joyID).Buttons And (2 ^ button)) <> 0
End Function

Public Function IsButtonPressed(ByVal joyID As Long, ByVal button As JoystickButton) As Boolean
    If Not ConnectedJoysticks(joyID) Then Exit Function
    Dim currentPressed As Boolean
    currentPressed = (JoystickStates(joyID).Buttons And (2 ^ button)) <> 0
    IsButtonPressed = currentPressed And Not ((LastJoystickStates(joyID).Buttons And (2 ^ button)) <> 0)
End Function

Public Function IsButtonReleased(ByVal joyID As Long, ByVal button As JoystickButton) As Boolean
    If Not ConnectedJoysticks(joyID) Then Exit Function
    Dim previouslyPressed As Boolean
    previouslyPressed = (LastJoystickStates(joyID).Buttons And (2 ^ button)) <> 0
    IsButtonReleased = Not ((JoystickStates(joyID).Buttons And (2 ^ button)) <> 0) And previouslyPressed
End Function

Private Function GetRawAxisValue(ByVal joyID As Long, ByVal axis As JoystickAxis) As Double
    If Not ConnectedJoysticks(joyID) Then
        GetRawAxisValue = DEFAULT_AXIS_CENTER
        Exit Function
    End If

    Select Case axis
        Case JoystickAxis.LEFT_ANALOG_X: GetRawAxisValue = JoystickStates(joyID).XPosition
        Case JoystickAxis.LEFT_ANALOG_Y: GetRawAxisValue = JoystickStates(joyID).YPosition
        Case JoystickAxis.RIGHT_ANALOG_X: GetRawAxisValue = JoystickStates(joyID).ZPosition
        Case JoystickAxis.RIGHT_ANALOG_Y: GetRawAxisValue = JoystickStates(joyID).RPosition
        Case JoystickAxis.TRIGGER_LT: GetRawAxisValue = (JoystickStates(joyID).VPosition / 65535)
        Case JoystickAxis.TRIGGER_RT: GetRawAxisValue = (JoystickStates(joyID).UPosition / 65535)
        Case Else: GetRawAxisValue = DEFAULT_AXIS_CENTER
    End Select
End Function

Private Function GetNormalizedAxisValue(ByVal joyID As Long, ByVal axis As JoystickAxis) As Double
    Dim rawValue As Long
    If Not ConnectedJoysticks(joyID) Then
        GetNormalizedAxisValue = 0
        Exit Function
    End If

    rawValue = GetRawAxisValue(joyID, axis)
    If axis = JoystickAxis.TRIGGER_LT Or axis = JoystickAxis.TRIGGER_RT Then
        GetNormalizedAxisValue = rawValue / 65535
    Else
        GetNormalizedAxisValue = (rawValue - DEFAULT_AXIS_CENTER) / DEFAULT_AXIS_CENTER
    End If
End Function

Public Function GetAxisValue(ByVal joyID As Long, ByVal axis As JoystickAxis, Optional ByVal normalized As Boolean = False) As Double
    If Not normalized Then
        GetAxisValue = GetRawAxisValue(joyID, axis)
    Else
        GetAxisValue = GetNormalizedAxisValue(joyID, axis)
    End If
End Function

Public Function IsMoving(ByVal joyID As Long, ByVal axis As JoystickAxis) As Boolean
    If Not ConnectedJoysticks(joyID) Then Exit Function
    Dim value As Long
    
    value = GetRawAxisValue(joyID, axis)
    IsMoving = Abs(value - DEFAULT_AXIS_CENTER) > Joysticks.Devices(joyID).Deadzone
End Function

Public Function IsDPadPressed(ByVal joyID As Long) As Boolean
    Dim dPadValue As DPadDirection
    dPadValue = AxisCore.GetDPadDirection(joyID)

    If dPadValue <> LastDPadState And dPadValue <> DPadDirection.CENTERED Then
        LastDPadState = dPadValue
        IsDPadPressed = True
    Else
        IsDPadPressed = False
    End If
End Function

Public Function GetDPadDirection(ByVal joyID As Long) As DPadDirection
    If Not ConnectedJoysticks(joyID) Then Exit Function

    Dim povValue As Long
    povValue = JoystickStates(joyID).POV
    
    If povValue = POV_CENTERED Or (povValue >= 0 And povValue <= 35999) Then
        GetDPadDirection = povValue
    Else
        GetDPadDirection = DPadDirection.CENTERED
    End If
End Function

Public Function GetAnalogStickDirection(ByVal joyID As Long, ByVal axisX As JoystickAxis, ByVal axisY As JoystickAxis) As JoystickAxisDirection
    If Not ConnectedJoysticks(joyID) Then
        GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_CENTERED
        Exit Function
    End If
    
    Dim xValue As Long
    Dim yValue As Long
    
    xValue = GetRawAxisValue(joyID, axisX) - DEFAULT_AXIS_CENTER
    yValue = GetRawAxisValue(joyID, axisY) - DEFAULT_AXIS_CENTER
    
    If Abs(xValue) < 2000 And Abs(yValue) < 2000 Then
        GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_CENTERED
    ElseIf Abs(xValue) > Abs(yValue) Then
        If xValue > 0 Then
            If yValue > 0 Then
                GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_DOWN_RIGHT
            ElseIf yValue < 0 Then
                GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_UP_RIGHT
            Else
                GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_RIGHT
            End If
        Else
            If yValue > 0 Then
                GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_DOWN_LEFT
            ElseIf yValue < 0 Then
                GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_UP_LEFT
            Else
                GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_LEFT
            End If
        End If
    Else
        If yValue > 0 Then
            GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_DOWN
        ElseIf yValue < 0 Then
            GetAnalogStickDirection = JoystickAxisDirection.AXIS_DIRECTION_UP
        End If
    End If
End Function
