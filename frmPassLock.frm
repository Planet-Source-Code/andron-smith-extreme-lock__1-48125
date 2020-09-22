VERSION 5.00
Begin VB.Form frmPassLock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmPassLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Username"
      Top             =   840
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   360
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Password"
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "frmPassLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DynPass As Recordset
Private DynConn As Connection


Private Sub Redraw_Form()
Me.Height = Screen.Height
Me.Width = Screen.Width
Me.Top = 0
Me.Left = 0
   txtUser.Top = Me.ScaleHeight / 2 - txtUser.Height / 2
   txtUser.Left = Me.ScaleWidth / 2 - txtUser.Width / 2
   
txtPassword.Top = (txtUser.Top + txtUser.Height) + 255
txtPassword.Left = txtUser.Left

End Sub
Private Function FindUser(ByVal UName As String) As String
On Error Resume Next
DynPass.Find "strUser = '" & UName & "'"
If DynPass.EOF = True Or DynPass.BOF = True Then
   DynPass.MoveFirst
   DynPass.AddNew
   DynPass!strUser = UName
   DynPass!strPass = txtPassword.Text
   DynPass.Update
   MsgBox "New Password Set"
   FindUser = txtPassword
   
   
   
   Exit Function
   Else
   FindUser = DynPass!strPass
End If


If Err > 0 Then MsgBox "Error In database: " & vbCrLf & Error




End Function
Private Function LoadDbase() As Boolean
On Error Resume Next
' Loads the Database
Set DynConn = New Connection ' note instead  of Entering Dim names as new object, it's good practice to Declare _
Dim on the top of the line. With out the new statement. And use Set Dclared = new Object
'Ex. :
'Dim Connections as connection
' Set Connectiosn = new connections.


Set DynPass = New Recordset

DynConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=ExUsers.mdb;"
DynConn.Open

DynPass.Open "Select * from DatAccount", DynConn, adOpenDynamic, adLockPessimistic

If DynConn.State <> adStateOpen Then LoadDbase = False
If DynPass.State <> adStateOpen Then LoadDbase = False
If DynConn.State = adStateOpen And DynPass.State = adStateOpen Then LoadDbase = True

End Function


Private Sub Form_Load()
X = LoadDbase()
If X <> True Then MsgBox "Error: In loading USername and Password Data -  " & Error

lngI = SetFocuses(Me.hwnd)



End Sub

Private Sub Form_Resize()
Redraw_Form


End Sub

Private Sub Form_Unload(Cancel As Integer)
' set DynConn = nothing
' Set Dynpass = nothing

End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If FindUser(txtUser.Text) = "NOTSET" Then
MsgBox "Username not found"
End

End If
  
     If txtPassword.Text = FindUser(txtUser.Text) Then
     End
     Else
     MsgBox "Username / Password", vbApplicationModal
     End If
     
End If



End Sub
