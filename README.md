Sub main() 
MsgBox ("從1加至50＝" & sum( 50 ) ) 
End Sub 
Function sum(ByVal x) As Long 
If x = 1 Then 
sum = 1 
Else 
sum = x + sum(x - 1) 
End If 
End Function
