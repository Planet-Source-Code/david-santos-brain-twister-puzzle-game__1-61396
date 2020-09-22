Attribute VB_Name = "modAPI"

Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, lpName As Any, lpType As Any) As Long
Public Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function CreateIconFromResource Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const LOAD_LIBRARY_AS_DATAFILE = &H2&

Public Const RT_RCDATA = 10&
Public Const RT_CURSOR = 1&
Public Const RT_ANICURSOR = 21&

Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type PuzzlePoint
    dist As Single
    angle As Single
End Type

Public Type PuzzlePiece
    centerx As Single
    centery As Single
    offsetx As Single
    offsety As Single
    offseta As Single
    angle As Single
    center As POINTAPI
    pointcount As Long
    points() As POINTAPI
    pinfo() As PuzzlePoint
    hrgn As Long
    selected As Boolean
    current As Boolean
    flipped As Integer
    snapped As Boolean
End Type

'IDR_101                 CURSOR  DISCARDABLE     "hand_closed.cur"
'IDR_102                 CURSOR  DISCARDABLE     "hand_open.cur"
'IDR_103                 CURSOR  DISCARDABLE     "hand_point.cur"
'IDR_104                 CURSOR  DISCARDABLE     "rotate.cur"
'IDR_105                 CURSOR  DISCARDABLE     "rotbot.cur"
'IDR_106                 CURSOR  DISCARDABLE     "rotbotleft.cur"
'IDR_107                 CURSOR  DISCARDABLE     "rotbotright.cur"
'IDR_108                 CURSOR  DISCARDABLE     "rotleft.cur"
'IDR_109                 CURSOR  DISCARDABLE     "rotright.cur"
'IDR_110                 CURSOR  DISCARDABLE     "rottop.cur"
'IDR_111                 CURSOR  DISCARDABLE     "rottopleft.cur"
'IDR_112                 CURSOR  DISCARDABLE     "rottopright.cur"

Public Enum CURSOR_TYPES
    CUR_HAND_CLOSED = 1
    CUR_HAND_OPEN = 2
    CUR_HAND_POINT = 3
    CUR_ROTATE = 4
    CUR_ROT_BOTTOM = 5
    CUR_ROT_BOTTOMLEFT = 6
    CUR_ROT_BOTTOMRIGHT = 7
    CUR_ROT_LEFT = 8
    CUR_ROT_RIGHT = 9
    CUR_ROT_TOPLEFT = 11
    CUR_ROT_TOP = 10
    CUR_ROT_TOPRIGHT = 12
End Enum

Public Type Puzzle
    picture As String
    solutions As Integer
    solution() As String
    solved As Boolean
    besttime As Long
End Type

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hrgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hrgn As Long, ByVal hBrush As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

Public Const ALTERNATE = 1
Public Const WINDING = 2
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020    ' (DWORD) dest = source
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

