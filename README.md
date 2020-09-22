<div align="center">

## Check if a string is an email address


</div>

### Description

This code returns a boolean expression that declares if a string is a valid email address or not. It returns true if the string is valid, false if not
 
### More Info
 
email as string

CheckIfEmail as boolean

it doesn't checks if the string actually is an email address, only if it is a valid email address.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Gabrielsen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-gabrielsen.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-gabrielsen-check-if-a-string-is-an-email-address__1-8008/archive/master.zip)





### Source Code

```
Public Function checkIfEmail(email As String) As Boolean
  Dim i As Integer
  Dim char As String
  Dim c() As String
  'checks if the string has the standard email pattern:
  If Not email Like "*@*.*" Then
   checkIfEmail = False
   Exit Function
  End If
  'splits the email-string with a "." delimeter and returns the subtring in the c-string array
  c = Split(email, ".", -1, vbBinaryCompare)
  'checks if the last substring has a length of either 2 or 3
  If Not Len(c(UBound(c))) = 3 And Not Len(c(UBound(c))) = 2 Then
   checkIfEmail = False
   Exit Function
  End If
  'steps through the last substring to see if it contains anything else unless characters from a to z
  For i = 1 To Len(c(UBound(c))) Step 1
   char = Mid(c(UBound(c)), i, 1)
   If Not (LCase(char) <= Chr(122)) Or Not (LCase(char) >= Chr(97)) Then
     checkIfEmail = False
     Exit Function
   End If
  Next i
  'steps through the whole email string to see if it contains any special characters:
  For i = 1 To Len(email) Step 1
   char = Mid(email, i, 1)
   If (LCase(char) <= Chr(122) And LCase(char) >= Chr(97)) _
     Or (char >= Chr(48) And char <= Chr(57)) _
     Or (char = ".") _
     Or (char = "@") _
     Or (char = "-") _
     Or (char = "_") Then
      checkIfEmail = True
   Else
     checkIfEmail = False
     Exit Function
   End If
  Next i
End Function
```

