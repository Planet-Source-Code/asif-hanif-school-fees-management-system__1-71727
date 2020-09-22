Attribute VB_Name = "modFx"
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  modFX                    #
'#    description :  Screen & GUI Effects     #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

'EXIT WINDOWS SETTINGS
Public Declare Function ExitWindows Lib "User32" (ByVal dwReturnCode As Long, ByVal uReserved As Integer) As Integer
Global Const EW_REBOOTSYSTEM = &H43
Global Const EW_RESTARTWINDOWS = &H42
Global Const EW_EXITWINDOWS = 0


' FOR RESOLUTION CHANGER
Private Declare Function EnumDisplaySettings Lib "User32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "User32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Private Const CCDEVICENAME = 32
    Private Const CCFORMNAME = 32
    Private Const DM_BITSPERPEL = &H60000
    Private Const DM_PELSWIDTH = &H80000
    Private Const DM_PELSHEIGHT = &H100000
Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    End Type

Public OldWidth1  As Integer
Public OldHeight1 As Integer
Public OldBPP1 As Integer


' EXIT WINDOWS FUNCTION
Function DoExitWindows()
    On Error Resume Next
    Dim RetVal As Integer
    RetVal = ExitWindows(EW_EXITWINDOWS, 0)
End Function

' FOR RESOLUTION VERIFIER
Function IsResolution(Width As Integer, Height As Integer) As Boolean
    If (Screen.Width / Screen.TwipsPerPixelX = Width) And (Screen.Height / Screen.TwipsPerPixelY = Height) Then
        IsResolution = True
    Else
        IsResolution = False
    End If
End Function

' FOR RESOLUTION CHANGER
Function ChangeRes(Width As Single, Height As Single, BPP As Integer) As Integer
    On Error GoTo ERROR_HANDLER
    Dim DevM As DEVMODE, i As Integer, ReturnVal As Boolean, _
    RetValue, OldWidth As Single, OldHeight As Single, _
    OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
    
    OldWidth1 = OldWidth
    OldHeight1 = OldHeight
    OldBPP1 = OldBPP
    
    i = 0
    Do
        ReturnVal = EnumDisplaySettings(0&, i, DevM)
        i = i + 1
    Loop Until (ReturnVal = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = Width
    DevM.dmPelsHeight = Height
    DevM.dmBitsPerPel = BPP
    Call ChangeDisplaySettings(DevM, 1)
    If RetValue = vbCancel Then
        DevM.dmPelsWidth = OldWidth
        DevM.dmPelsHeight = OldHeight
        DevM.dmBitsPerPel = OldBPP
        Call ChangeDisplaySettings(DevM, 1)
        ChangeRes = 0
    Else
        ChangeRes = 1
    End If
    Exit Function
ERROR_HANDLER:
    ChangeRes = 0
End Function


' EXIT APPLICATION FOR RESOLUTION CHANGER
Function ExitApplication(Width As Single, Height As Single, BPP As Integer) As Integer
    On Error GoTo ERROR_HANDLER
    Dim DevM As DEVMODE, i As Integer, ReturnVal As Boolean, _
    RetValue, OldWidth As Single, OldHeight As Single, _
    OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
    i = 0
    Do
        ReturnVal = EnumDisplaySettings(0&, i, DevM)
        i = i + 1
    Loop Until (ReturnVal = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = Width
    DevM.dmPelsHeight = Height
    DevM.dmBitsPerPel = BPP
    Call ChangeDisplaySettings(DevM, 1)
'    If RetValue = vbCancel Then
'        DevM.dmPelsWidth = OldWidth
'        DevM.dmPelsHeight = OldHeight
'        DevM.dmBitsPerPel = OldBPP
'        Call ChangeDisplaySettings(DevM, 1)
 '       ChangeRes = 0
 '   Else
 '       ChangeRes = 1
 '   End If
    Exit Function
ERROR_HANDLER:
    'ChangeRes = 0
End Function


