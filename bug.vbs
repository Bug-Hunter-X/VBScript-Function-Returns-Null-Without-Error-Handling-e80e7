Function MyFunc(param1)
  If IsEmpty(param1) Then
    MyFunc = Null ' This will cause an error if MyFunc is used later without checking for Null
  Else
    MyFunc = param1 * 2
  End If
End Function

Dim result
result = MyFunc(5)
MsgBox result

result = MyFunc(Empty)
MsgBox result 'Error occurs here because Null is not handled properly 