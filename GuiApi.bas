Attribute VB_Name = "GuiApi"
Option Explicit
Option Compare Text

Public Const Filters = "All program files" & vbNullChar & "*.php;*.htm;*.html;*.svg" & vbNullChar & _
                       "php files" & vbNullChar & "*.php" & vbNullChar & _
                       "html files" & vbNullChar & "*.htm;*.html" & vbNullChar & _
                       "svg files" & vbNullChar & "*.svg" & vbNullChar & _
                       "All files" & vbNullChar & "*.*"

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Private Const OFN_LONGNAMES = &H200000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_NODEREFERENCELINKS = &H100000

Private Type OPENFILENAME
    lStructSize As Long
    hInstance As Long
    hwndOwner As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
       Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
 "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&
Const INFINITE& = &HFFFF

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type SelectorMethod
    Min As Integer
    Max As Integer
    Current As Double
End Type

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal HDC As Long, lpRect As RECT) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal HWnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long

Public CheckedButton(20) As Boolean
Public CheckedButtonRect(20) As Boolean
Public CheckedButtonBrush(20) As Boolean
Public CheckedButtonLight(20) As Boolean

Public CurrentButton As Integer
Public CurrentButtonRect As Integer
Public CurrentButtonBrush As Integer
Public CurrentButtonLight As Integer

Public SelColorIndex As Integer

Public ClipBoardGotData As Boolean

Public SelMethod(10) As SelectorMethod









Function GetOpenName(Optional ByVal WindowTitle As String = "Load File", _
                     Optional ByVal Filters As String = Filters, _
                     Optional ByVal DefaultFileName As String = "")
 Dim ret As Long
 Dim DlgInfo As OPENFILENAME
 
 With DlgInfo
      .lStructSize = Len(DlgInfo)
      .hwndOwner = 0
      .lpstrFilter = Filters
      .nFilterIndex = 1
      .lpstrFile = DefaultFileName & Space$(1024) & vbNullChar & vbNullChar
      .nMaxFile = Len(.lpstrFile)
      .lpstrDefExt = vbNullChar & vbNullChar
      .lpstrFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      .nMaxFileTitle = Len(.lpstrFileTitle)
      .lpstrInitialDir = CurDir + vbNullChar
      .lpstrTitle = WindowTitle
      .Flags = OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS
 End With
  GetOpenName = GetOpenFileName(DlgInfo)
  GetOpenName = Left(DlgInfo.lpstrFile, InStr(DlgInfo.lpstrFile, vbNullChar) - 1)
End Function





