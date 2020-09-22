<div align="center">

## IsIP Function v2


</div>

### Description

Updated version of IsIP function. Should clear up all the problems in the last one. Here ya go, =)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel M](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-m.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-m-isip-function-v2__1-57999/archive/master.zip)





### Source Code

```
Private Function IsIP(strIP As String) As Boolean
 Dim splitIP() As String, i As Long
 IsIP = True 'Starts out as true
 splitIP$ = Split(strIP$, ".", -1, 1) 'Split IP To check value
 '========================================
 'Things we must check to verify IP
 '1. Make sure there are 4 sections to IP
 '2. Make sure each section of IP is not
 ' greater than 255
 '3. Make sure each section of IP does
 ' not t contain a negative
 '4. Make sure each section of IP is nume-
 ' ric
 '5. Make sure first section of IP is not
 ' 0
 '=======================================
 If UBound(splitIP$) <> 3 Then
  IsIP = False 'make sure there is only 4 nodes =)
 Else
  For i = 0 To UBound(splitIP$) 'loop through array and check 3 things
   If IsNumeric(splitIP(i)) = False Then
    IsIP = False
    Exit For
   Else
    If splitIP(0) = 0 Then 'first digit cannot be 0
     IsIP = False
     Exit For
    End If
    If splitIP(i) > 255 Or splitIP(i) < 0 Then
     IsIP = False
     Exit For
    End If
   End If
  Next i
 End If
End Function
```

