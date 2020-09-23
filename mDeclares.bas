Attribute VB_Name = "mDeclares"
Option Explicit

' I usually put constants, types, and functions in the forms/modules/classes
' that call them, but this project had so many (and the potential to have
' even more), I decided to put them all in one module.

' Constants.
Public Const CCDEVICENAME = 32&
Public Const CCFORMNAME = 32&
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000

Public Const GW_CHILD = 5&
Public Const GWL_WNDPROC = -4&
Public Const GWL_HINSTANCE = -6&
Public Const GWL_HWNDPARENT = -8&
Public Const GWL_STYLE = -16&
Public Const GWL_EXSTYLE = -20&
Public Const GWL_USERDATA = -21&
Public Const GWL_ID = -12&
Public Const WS_CHILD = &H40000000
Public Const SWP_NOMOVE = 2&
Public Const SWP_NOSIZE = 1&
Public Const FLAGS = SWP_NOMOVE& Or SWP_NOSIZE&
Public Const HWND_TOPMOST = -1&
Public Const HWND_NOTOPMOST = -2&
Public Const SPI_SCREENSAVERRUNNING = 97&

Public Const Pi       As Single = 3.14159265358979
Public Const PiDiv180 As Single = 1.74532925199433E-02
Public Const TPi      As Single = 6.28318530717959

Public Const HORZRES = 8&
Public Const VERTRES = 10&

' Types.
Public Type tPoint
    X As Long
    Y As Long
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type RGBQUAD
    b As Byte
    g As Byte
    r As Byte
    a As Byte
End Type

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Public Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Public Type DEVMODE
    dmDeviceName    As String * CCDEVICENAME
    dmSpecVersion   As Integer
    dmDriverVersion As Integer
    dmSize          As Integer
    dmDriverExtra   As Integer

    dmFields        As Long
    dmOrientation   As Integer
    dmPaperSize     As Integer
    dmPaperLength   As Integer
    dmPaperWidth    As Integer
    dmScale         As Integer
    dmCopies        As Integer
    dmDefaultSource As Integer
    dmPrintQuality  As Integer
    dmColor         As Integer
    dmDuplex        As Integer
    dmYResolution   As Integer
    dmTTOption      As Integer
    dmCollate       As Integer

    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Public Type udtFrameRate
    Text  As String
    Value As Long
    Ticks As Long
End Type

' Used to read/write blur settings in INI.
Public Type udtBlurSettings
    Enabled  As Long
    Quick    As Long
    Strength As Long
End Type

' Used to read/write buffer settings in INI.
Public Type udtBufferSettings
    Width            As Long
    Height           As Long
    ScreenIndex      As Long
    BackColor        As Long
    FillColor        As Long
    StretchMode      As Long
    DisplayFrameRate As Long
End Type

' Used to read/write flower settings in INI.
Public Type udtFlowerSettings
    FlowerCount     As Long
    PetalCount      As Long
    PetalWidth      As Long
    PetalPointiness As Long
End Type

Public Type udtSettings
    Mode   As String
    Flower As udtFlowerSettings
    Buffer As udtBufferSettings
    Blur   As udtBlurSettings
End Type

Public Type udtFlowerColorPart
    Value     As Integer
    Direction As Integer
End Type

Public Type udtFlowerColor
    r As udtFlowerColorPart
    g As udtFlowerColorPart
    b As udtFlowerColorPart
End Type

' Used to render the flowers.
Public Type udtFlower

    Points(3)   As tPoint
    Center      As tPoint
    Direction   As tPoint ' not really, but 2 related values
    
    Bounce      As Single
    BounceRate  As Single
    
    Color       As udtFlowerColor
    
    Angle       As Long
    Spin        As Long
    Pointiness  As Long
    
    PetalCount  As Long
    PetalWidth  As Long
    PetalHeight As Long
    
End Type

' API Functions/Subs
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As tPoint, ByVal nCount As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

