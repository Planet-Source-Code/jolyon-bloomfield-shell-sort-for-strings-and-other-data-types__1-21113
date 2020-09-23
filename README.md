<div align="center">

## Shell Sort for Strings and other data types


</div>

### Description

Shell sort routine, created for strings, but easily changed for any data type. Pass in an array. Features arguments for the last element of the array to be sotred, and also the first. Fast sorting routine, should be compatible with all versions of VB, because it is straight maths, although the "optional" keyword in the declaration isn't compatible with earlier versions. Enjoy
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jolyon Bloomfield](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jolyon-bloomfield.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jolyon-bloomfield-shell-sort-for-strings-and-other-data-types__1-21113/archive/master.zip)





### Source Code

```
Public Sub Sort(ByRef SortArray() As String, ByVal MaxRow As Integer, Optional ByVal MinRow As Integer = 1)
' Does a shell sort - fairly fast, and flexible
' In this case, sorts strings, but can easily be modified
' To suit other data types - simply change the definition of SortArray()
' and the next line, to the data type of your choice.
Dim TempSwap As String
Dim Offset As Integer
Dim Switch As Integer
Dim Limit As Integer
Dim Row As Integer
' Set comparison offset to half the number of records in SortArray:
Offset = (MaxRow - MinRow + 1) \ 2
Do While Offset > 0     ' Loop until offset gets to zero.
 Limit = MaxRow - Offset
 Do
  Switch = 0     ' Assume no switches at this offset.
  ' Compare elements and switch ones out of order:
  For Row = MinRow To Limit
   If UCase(SortArray(Row)) > UCase(SortArray(Row + Offset)) = True Then
    TempSwap = SortArray(Row)
    SortArray(Row) = SortArray(Row + Offset)
    SortArray(Row + Offset) = TempSwap
    Switch = Row
   End If
  Next Row
  ' Sort on next pass only to where last switch was made:
  Limit = Switch - Offset
 Loop While Switch
 ' No switches at last offset, try one half as big:
 Offset = Offset \ 2
Loop
End Sub
```

