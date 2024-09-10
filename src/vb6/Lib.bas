Attribute VB_Name = "Lib"
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

Public Sub DeleteINIKey(ByVal iniFile As String, ByVal section As String, ByVal key As String)
    WritePrivateProfileString section, key, vbNullString, iniFile
End Sub
