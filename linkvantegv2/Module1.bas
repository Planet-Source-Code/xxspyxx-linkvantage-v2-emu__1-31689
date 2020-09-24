Attribute VB_Name = "Module1"
Public Const URL = "http://www.linkvantage.com/signup.asp?referrer=XxSpyxX"
Public Const email = "Tweeqit2@hotmail.com"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1



Function GetCookie(Data As String)
Data = Right(Data, Len(Data) - InStr(Data, "Set-Cookie: ") - 11)
Data = Left(Data, InStr(Data, vbCrLf) - 1)
GetCookie = Data
End Function




Public Sub gotoweb()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub
