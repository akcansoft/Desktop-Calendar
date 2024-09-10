VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Define a private object to interact with the Windows Script Host (WSH)
Private wsh As Object

' Initialize the class by creating an instance of the WScript.Shell object
Private Sub Class_Initialize()
  Set wsh = CreateObject("WScript.Shell")
End Sub

' Clean up the class by releasing the WScript.Shell object
Private Sub Class_Terminate()
  Set wsh = Nothing
End Sub

' Function to read a value from the registry
' Parameters:
'   keyPath: The full path of the registry key to read from
' Returns:
'   The value of the registry key as a string, or an empty string if an error occurs
Public Function ReadKey(keyPath As String) As String
  On Error Resume Next
  ReadKey = wsh.RegRead(keyPath) ' Attempt to read the registry key value
  If Err.Number <> 0 Then ' Check if an error occurred
    ReadKey = "" ' Return an empty string if there was an error
  End If
  On Error GoTo 0 ' Reset error handling
End Function

' Function to write a value to the registry
' Parameters:
'   keyPath: The full path of the registry key to write to
'   value: The value to write to the registry key
' Returns:
'   True if the value was successfully written, False otherwise
Public Function WriteKey(keyPath As String, ByVal value As String) As Boolean
  On Error Resume Next
  wsh.RegWrite keyPath, value ' Attempt to write the value to the registry key
  If Err.Number = 0 Then ' Check if the write operation was successful
    WriteKey = True ' Return True if successful
  Else
    WriteKey = False ' Return False if an error occurred
  End If
  On Error GoTo 0 ' Reset error handling
End Function

' Function to delete a registry key
' Parameters:
'   keyPath: The full path of the registry key to delete
' Returns:
'   True if the key was successfully deleted, False otherwise
Public Function DeleteKey(keyPath As String) As Boolean
  On Error Resume Next
  wsh.RegDelete keyPath ' Attempt to delete the registry key
  If Err.Number = 0 Then ' Check if the delete operation was successful
    DeleteKey = True ' Return True if successful
  Else
    DeleteKey = False ' Return False if an error occurred
  End If
  On Error GoTo 0 ' Reset error handling
End Function

