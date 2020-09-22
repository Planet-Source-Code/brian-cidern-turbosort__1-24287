<div align="center">

## TurboSort


</div>

### Description

<p>

Sort arrays much faster with a better string

swapping routine!

</p>

<p>

Wow, I couldn't believe all the rewrites of the

same sorting routines in PSC. "Look at

mine", "No, use mine", yadda, yadda, yadda. They

all use the horribly slow:<br>

<pre>

vTemp = String1

String1 = String2

String1 = vTemp

</pre>

</p>

<p>

Geezzzz - When you have to sort 30,000+ strings

this is slllooooowwwwwww.

</p>

<p>

Here's a solution. It uses the the same sorting

routine (or choose your own), but implements a much

faster swap routine using the CopyMemory() API. Now,

instead of swapping strings, which in my case could

be up to 9,000 characters, you are only swapping a

4 byte memory address.

</p>

<p>

Rock On!!

</p>
 
### More Info
 
Create a new EXE and throw in Command1 - Paste the rest.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Cidern](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-cidern.md)
**Level**          |Beginner
**User Rating**    |4.9 (69 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-cidern-turbosort__1-24287/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Sub CopyMemory _
 Lib "kernel32" _
 Alias "RtlMoveMemory" ( _
 lpDest As Any, _
 lpSource As Any, _
 ByVal cbCopy As Long _
 )
Private Sub Command1_Click()
 ' Sort an array with CopyMemory()
 Dim i As Integer
 Dim str_Unsorted As String, _
 str_Sorted As String
 ' Populate some sample data
 Dim vArray(25) As String
 vArray(0) = "EFGHIJKLMNOPQRSTUVWXYZABCD"
 vArray(1) = "RSTUVWXYZABCDEFGHIJKLMNOPQ"
 vArray(2) = "PQRSTUVWXYZABCDEFGHIJKLMNO"
 vArray(3) = "DEFGHIJKLMNOPQRSTUVWXYZABC"
 vArray(4) = "IJKLMNOPQRSTUVWXYZABCDEFGH"
 vArray(5) = "ZABCDEFGHIJKLMNOPQRSTUVWXY"
 vArray(6) = "HIJKLMNOPQRSTUVWXYZABCDEFG"
 vArray(7) = "LMNOPQRSTUVWXYZABCDEFGHIJK"
 vArray(8) = "STUVWXYZABCDEFGHIJKLMNOPQR"
 vArray(9) = "TUVWXYZABCDEFGHIJKLMNOPQRS"
 vArray(10) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
 vArray(11) = "CDEFGHIJKLMNOPQRSTUVWXYZAB"
 vArray(12) = "VWXYZABCDEFGHIJKLMNOPQRSTU"
 vArray(13) = "MNOPQRSTUVWXYZABCDEFGHIJKL"
 vArray(14) = "FGHIJKLMNOPQRSTUVWXYZABCDE"
 vArray(15) = "JKLMNOPQRSTUVWXYZABCDEFGHI"
 vArray(16) = "YZABCDEFGHIJKLMNOPQRSTUVWX"
 vArray(17) = "XYZABCDEFGHIJKLMNOPQRSTUVW"
 vArray(18) = "OPQRSTUVWXYZABCDEFGHIJKLMN"
 vArray(19) = "BCDEFGHIJKLMNOPQRSTUVWXYZA"
 vArray(20) = "GHIJKLMNOPQRSTUVWXYZABCDEF"
 vArray(21) = "KLMNOPQRSTUVWXYZABCDEFGHIJ"
 vArray(22) = "NOPQRSTUVWXYZABCDEFGHIJKLM"
 vArray(23) = "WXYZABCDEFGHIJKLMNOPQRSTUV"
 vArray(24) = "QRSTUVWXYZABCDEFGHIJKLMNOP"
 vArray(25) = "UVWXYZABCDEFGHIJKLMNOPQRST"
 ' Here's the unsorted array
 For i = 0 To UBound(vArray)
 str_Unsorted = str_Unsorted & vArray(i) & vbCrLf
 Next i
 MsgBox str_Unsorted
 ' Sort the array
 SortMe vArray
 ' Here's the sorted array
 For i = 0 To UBound(vArray)
 str_Sorted = str_Sorted & vArray(i) & vbCrLf
 Next i
 MsgBox str_Sorted
End Sub
Sub SortMe(varArray() As String)
 Dim i As Long, j As Long
 Dim l_Count As Long
 Dim l_Hold As Long
 ' Typical sorting routine
 l_Count = UBound(varArray)
 For i = 0 To l_Count
 For j = i + 1 To l_Count
 If varArray(i) > varArray(j) Then
 ' Here's the juice!
 SwapStrings varArray(i), varArray(j)
 End If
 Next
 Next
End Sub
Sub SwapStrings(pbString1 As String, pbString2 As String)
 Dim l_Hold As Long
 CopyMemory l_Hold, ByVal VarPtr(pbString1), 4
 CopyMemory ByVal VarPtr(pbString1), ByVal VarPtr(pbString2), 4
 CopyMemory ByVal VarPtr(pbString2), l_Hold, 4
End Sub
```

