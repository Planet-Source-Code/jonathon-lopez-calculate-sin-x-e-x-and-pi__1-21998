<div align="center">

## Calculate Sin\(x\), e^\(x\) , and Pi


</div>

### Description

This code allows you to calculate Sin(x), e^(x), and Pi without using any of the built-in VB functions
 
### More Info
 
The way this code works is by using what is called Taylor Expansion. By generating a taylor series using derivatives it is possible to take the sum of the elements in that series to find such functions as Sin(x) and even Pi.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathon Lopez](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathon-lopez.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathon-lopez-calculate-sin-x-e-x-and-pi__1-21998/archive/master.zip)





### Source Code

```
'Sin(x) function
'Note: this is in radians, not degrees
Public Function Sine(x as Double) as Double
Dim i As Integer, sum As Double: sum = 0
'Calculate the taylor expansion of sin
For i = 1 To 10
  sum = sum + (((-1) ^ (i + 1)) * ((x) ^ (2 * i - 1)) / fact(2 * i - 1))
Next i
Sine=sum
End Function
'e^(x) function
Public Function e(x as Integer) as Double
Dim i As Integer, sum As Double: sum = 0
'Calculate the Taylor expansion of e
For i = 0 To 150
  sum = sum + (x ^ i) / fact(i)
Next i
e=sum
End Function
'Pi function
Public Function pi() as Double
Dim i As Integer, sum As Double: sum = 0
For i = 1 To 15000
  sum = sum + ((-1) ^ (i + 1)) * (1 ^ (2 * i - 1)) / (2 * i - 1)
Next i
pi = sum * 4
End Function
'Function that calculates factorials
Public Function fact(n As Integer) As Double
Dim i As Long, r As Double: r = 1
If n = 0 Then fact = 1
For i = 1 To n
  r = i * r
Next i
fact = r
End Function
```

