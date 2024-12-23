Function MyFunc(param1)
  If IsEmpty(param1) Then
    MyFunc = 0 ' Return 0 or a default value instead of Null
  Else
    MyFunc = param1 * 2
  End If
End Function

Dim result
result = MyFunc(5)
MsgBox result

result = MyFunc(Empty)
MsgBox result 'Now handles the case where the input is empty without an error