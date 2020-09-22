Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Const READ_CONTROL As Long = &H20000
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)
Public Const SYNCHRONIZE As Long = &H100000
Public Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Public Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Const ERROR_SUCCESS As Long = 0&
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const REG_SZ As Long = 1

Public Const LWA_COLORKEY As Integer = &H1
Public Const LWA_ALPHA As Integer = &H2
Public Const GWL_EXSTYLE As Integer = (-20)
Public Const WS_EX_LAYERED As Long = &H80000

Public Const LF_FACESIZE As Long = 32
Public FontDialog As CHOOSEFONTS
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long

Public Const CF_SCREENFONTS As Long = &H1
Public Const CF_EFFECTS As Long = &H100&
Public Const CF_INITTOLOGFONTSTRUCT As Long = &H40&
Public Const WH_CBT As Long = 5

Public Const GWL_HINSTANCE As Long = (-6)
Public Const CC_RGBINIT As Long = &H1
Public Const CC_FULLOPEN As Long = &H2
Public Const CC_PREVENTFULLOPEN As Long = &H4
Public Const CC_SHOWHELP As Long = &H8
Public Const CC_ENABLEHOOK As Long = &H10
Public Const CC_ENABLETEMPLATE As Long = &H20
Public Const CC_ENABLETEMPLATEHANDLE As Long = &H40
Public Const CC_SOLIDCOLOR As Long = &H80
Public Const CC_ANYCOLOR As Long = &H100
Public Const COLOR_FLAGS As Long = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Public Type CHOOSEFONTS
    lStructSize As Long
    hwndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public m_IgnoreEvents As Boolean
Private ParenthWnd As Long

Public Function ShowFont(ByVal hWnd As Long, ByVal startingFontName As String, Optional ByVal centerForm As Boolean = True) As SelectedFont

  Dim ret As Long
  Dim lfLogFont As LOGFONT
  Dim hInst As Long
  Dim Thread As Long
  Dim i As Integer
  Static ultimo_tamanho As Long
  Static ultimo_italic As Long
  Static ultimo_weight As Long

    ParenthWnd = hWnd
    FontDialog.nSizeMax = 0
    FontDialog.nSizeMin = 0
    FontDialog.nFontType = Screen.FontCount
    FontDialog.hwndOwner = hWnd
    FontDialog.hdc = 0
    FontDialog.lpfnHook = 0
    FontDialog.lCustData = 0
    lfLogFont.lfHeight = ultimo_tamanho
    lfLogFont.lfItalic = ultimo_italic
    lfLogFont.lfWeight = ultimo_weight
    FontDialog.lpLogFont = VarPtr(lfLogFont)
    If FontDialog.iPointSize = 0 Then
        FontDialog.iPointSize = 10 * 10
    End If

    FontDialog.lpTemplateName = Space$(2048)

    FontDialog.lStructSize = Len(FontDialog)

    If FontDialog.flags = 0 Then
        FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    End If

    For i = 0 To Len(startingFontName) - 1
        lfLogFont.lfFaceName(i) = Asc(Mid(startingFontName, i + 1, 1))
    Next i

    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()

    ret = ChooseFont(FontDialog)

    If ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
        Next i

        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
        ultimo_tamanho = lfLogFont.lfHeight
        ultimo_italic = lfLogFont.lfItalic
        ultimo_weight = lfLogFont.lfWeight
        Exit Function
      Else
        ShowFont.bCanceled = True
        Exit Function
    End If

End Function

Public Function WillRunAtStartup(ByVal App_Name As String) As Boolean

  Dim hKey As Long
  Dim value_type As Long

    If RegOpenKeyEx(HKEY_CURRENT_USER, _
                    "Software\Microsoft\Windows\CurrentVersion\Run", _
                    0, KEY_READ, hKey) = ERROR_SUCCESS _
                    Then
        WillRunAtStartup = _
                           (RegQueryValueEx(hKey, App_Name, _
                           ByVal 0&, value_type, ByVal 0&, ByVal 0&) = _
                           ERROR_SUCCESS)

        RegCloseKey hKey
      Else
        WillRunAtStartup = False
    End If

End Function

Public Sub SetRunAtStartup(ByVal App_Name As String, ByVal app_path As String, Optional ByVal run_at_startup As Boolean = True)

  Dim hKey As Long
  Dim key_value As String
  Dim status As Long

    On Error GoTo SetStartupError

    If RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&) <> ERROR_SUCCESS Then
        MsgBox "Error " & Err.Number & " opening key" & vbCrLf & Err.Description
        Exit Sub
    End If

    If run_at_startup Then

        key_value = app_path & "\" & App_Name & ".exe" & vbNullChar
        status = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, _
                 ByVal key_value, Len(key_value))
        If status <> ERROR_SUCCESS Then
            MsgBox "Error " & Err.Number & " setting key" & vbCrLf & Err.Description
        End If

      Else

        RegDeleteValue hKey, App_Name

    End If

    RegCloseKey hKey

exit_sub:
Exit Sub

SetStartupError:
    MsgBox Err.Number & " " & Err.Description
    Resume exit_sub
    
End Sub
