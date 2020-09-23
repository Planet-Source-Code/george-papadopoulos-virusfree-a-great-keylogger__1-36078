Attribute VB_Name = "Module1"
Global LastKey As String
Global timeout As Byte
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Public Function GetCapslock() As Boolean
' Return or set the Capslock toggle.

GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)

End Function

Public Function GetShift() As Boolean

' Return or set the Capslock toggle.

GetShift = CBool(GetAsyncKeyState(vbKeyShift))

End Function

