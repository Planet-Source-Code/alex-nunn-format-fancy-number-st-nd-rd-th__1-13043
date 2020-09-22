<div align="center">

## Format Fancy Number \(\#\#st, \#\#nd, \#\#rd, \#\#th\)


</div>

### Description

This function adds st, nd, rd, or th to the end of a string of numbers based on what the number is. For example, the following code can produce the following output, "Thursday, November 23rd, 2000" : Format(Date, "dddd, mmmm ") & FormatFancyNumber(Day(Date)) & ", " & Year(Date)
 
### More Info
 
All this function needs is an integer number in string format.

It returns the number in string format with st, nd, rd, or th added to the end as needed.

The code currently only works with integer size numbers. This should be easy to change though considering the size of code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alex Nunn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex-nunn.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alex-nunn-format-fancy-number-st-nd-rd-th__1-13043/archive/master.zip)





### Source Code

```
Public Function FormatFancyNumber(ByVal sNumber As String) As String
 Dim iTemp As Integer
 iTemp = Int(sNumber)
 If 4 < iTemp And iTemp < 20 Then
  FormatFancyNumber = sNumber & "th"
 Else
  Select Case iTemp Mod 10
   Case 1
    FormatFancyNumber = sNumber & "st"
   Case 2
    FormatFancyNumber = sNumber & "nd"
   Case 3
    FormatFancyNumber = sNumber & "rd"
   Case Else
    FormatFancyNumber = sNumber & "th"
  End Select
 End If
End Function
```

