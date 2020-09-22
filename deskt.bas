Attribute VB_Name = "Module1"
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetFocuses Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

' This is a password Program... It uses JET Engine to add your password to a Database
' If the user Is running the program for the first time... A New password
' is set, if the user is returning... The password is tested..
'
'


Public Sub Main()
  Dim DesktopdC As Long
  Dim strName As String
  Dim lngBuffer As Long
  strName = String$(255, 0)
  DesktopdC = GetDesktopWindow
  lngBuffer = GetUserName(strName, Len(strName))

  
  Load frmPassLock
 frmPassLock.txtUser.Text = strName
 frmPassLock.Show
 
  
End Sub
