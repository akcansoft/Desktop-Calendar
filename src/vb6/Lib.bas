Attribute VB_Name = "Lib"
Private Const REG_SZ As Long = 1
'Const HKEY_CURRENT_USER = &H80000001

Private Const ERROR_NONE = 0
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
  ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long
  
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
  ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
  ByVal lpData As String, lpcbData As Long) As Long
  
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
  ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
  ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
  ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
  ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
  ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
  ByVal lpFileName As String) As Long

Public Function ReadINI(ByVal iniFile As String, ByVal section As String, _
  ByVal key As String, Optional ByVal defaultValue As String = "") As String
  
  Dim buffer As String * 255
  Dim returnLength As Long
  
  returnLength = GetPrivateProfileString(section, key, "", buffer, Len(buffer), App.Path & "\" & iniFile)
  
  If returnLength > 0 Then
      ReadINI = Left(buffer, returnLength)
  Else
      ReadINI = defaultValue
  End If
End Function

Public Sub WriteINI(ByVal iniFile As String, ByVal section As String, _
  ByVal key As String, ByVal value As String)
  WritePrivateProfileString section, key, value, App.Path & "\" & iniFile
End Sub

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long, lrc As Long, lType As Long
    Dim sValue As String
    On Error GoTo QueryValueExError
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        Case Else
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
       Dim lRetVal As Long
       Dim hKey As Long
       Dim vValue As Variant

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       vValue = Replace(vValue, Chr$(0), "")
       QueryValue = vValue
       RegCloseKey (hKey)
End Function

