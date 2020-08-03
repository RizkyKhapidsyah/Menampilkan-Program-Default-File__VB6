Attribute VB_Name = "Module1"
Declare Function FindExecutable Lib "shell32.dll" _
Alias "FindExecutableA" (ByVal lpFile As String, _
ByVal lpDirectory As String, ByVal lpResult As _
String) As Long

Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


