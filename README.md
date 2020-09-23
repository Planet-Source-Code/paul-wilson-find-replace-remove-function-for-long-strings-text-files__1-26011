<div align="center">

## Find/Replace/Remove Function for long strings/text files


</div>

### Description

This Function Searches a user defined string for a user defined search criteria it will return the postion of matchs in an array also will replace or remove criteria from string
 
### More Info
 
'strData = Target String to search - This is required

'strValue = String to search for within strData - This is required

'strReplace = String to replace any matches found with - This is optional

'boolRemove = this is a switch to activate the remove feature - Optional

'boolReplace = this is a swith to activate the replace feature - Optional

'This functions purpose is to search for a user defined string within another user defined string

'It provides the ability to remove or replace that string

'The function returns an array with the start position of each match found. This feature means that the programer

'using this function has the ability to easily and quickly recreate the powerful find and replace features found in

'products like MS Word

'By recieving matches in an array you can quickly create a find first find next type feature as well

'THIS VERSION IS CASE SENSITIVE coming soon is the option for a case sensetive or generic search

'arrPos = This is a return the position of each match in an array

'lngFoundCount = This returns the count of matches found

'The actual function is a long Data Type and returns an error value

'0 = Search Completed No Results

'1 = Search Completed Matches Found

'2 = No Target String defined

'3 = No Search String Defined

'4 = No Replace String Defined when Replace is set to true

'5 = Both Replace and Remove features have been set to true

'6 = Unexpected Error


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Wilson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-wilson.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-wilson-find-replace-remove-function-for-long-strings-text-files__1-26011/archive/master.zip)





### Source Code

```
Public Function findInString(strdata As String, strValue As String, strReplace As String, arrPos() As Long, lngFoundCount As Long, boolRemove As Boolean, boolReplace As Boolean) As Long
'This functions purpose is to search for a user defined string within another user defined string
'It provides the ability to remove or replace that string
'The function returns an array with the start position of each match found. This feature means that the programer
'using this function has the ability to easily and quickly recreate the powerful find and replace features found in
'products like MS Word
'By recieving matches in an array you can quickly create a find first find next type feature as well
'THIS VERSION IS CASE SENSITIVE coming soon is the option for a case sensetive or generic search
'---------------------------------------
'Function Info
'Passed Variables
'strData = Target String to search - This is required
'strValue = String to search for within strData - This is required
'strReplace = String to replace any matches found with - This is optional
'arrPos = This is a return the position of each match in an array
'lngFoundCount = This returns the count of matches found
'boolRemove = this is a switch to activate the remove feature - Optional
'boolReplace = this is a swith to activate the replace feature - Optional
'The actual function is a long Data Type and returns an error value
'0 = Search Completed No Results
'1 = Search Completed Matches Found
'2 = No Target String defined
'3 = No Search String Defined
'4 = No Replace String Defined when Replace is set to true
'5 = Both Replace and Remove features have been set to true
'6 = Unexpected Error
'---------------------------------------
Dim arrByteTarget() As Byte ' Declare array to contain Target string
Dim arrByteFind() As Byte ' Declare array to contain Search String
Dim arrByteReplace() As Byte 'Declare array to contain Replace String
Dim arrByteTempLeft() As Byte 'Declare a temoporary array to contain data to the left of a match
Dim arrByteTempRight() As Byte 'Declare temporary array to contain data to the right of a match
Dim lngLoopTarget As Long 'Declare A Loop counter
Dim lngFindStart As Long 'Declare a start position container
Dim lngloopStep As Long 'Declare another loop counter
Dim lngStepTemp As Long 'Declare yet another loop Counter
Dim lngStrValBytCount As Long 'Declare Search String Byte Count container
Dim lngStrRemBytCount As Long 'Declare Replace String Byte Count Container
Dim lngTempBound As Long 'Declare a temporary Long Variable Container
Dim boolFoundTemp As Boolean 'Declare a temporary Found Switch
Dim boolFound As Boolean 'Declare a Found Switch
Dim boolSpaceAdded As Boolean 'Declare a Space has been added to front of string switch
On Error GoTo ErrorHandler 'Always be Wary for the unexpected
If Len(strdata) = 0 Then 'Check the target string has data
  findInString = 2
  GoTo ExitFunction
End If
If Len(strValue) = 0 Then 'check the search string has data
  findInString = 3
  GoTo ExitFunction
End If
If boolReplace = True Then 'check to see if replace has been selected
  If Len(strReplace) = 0 Then 'if it has check to see replace string has data
    findInString = 4
    GoTo ExitFunction
  Else
    strReplace = Chr$(32) & strReplace 'if it does add a space to the front of it for Padding
    arrByteReplace = strReplace
  End If
End If
If boolReplace = True And boolRemove = True Then 'check that both replace and remove arnt selected
  findInString = 5
  GoTo ExitFunction
End If
If Len(strValue) = 1 Then 'Check to see if the search value is a space if it is dont add spaces
  If Asc(strValue) = 32 Then
    boolSpaceAdded = False
    GoTo StartSearch
  End If
End If
strValue = Chr$(32) & strValue & Chr$(32) 'add spaces to front and back of search string this is to make sure it doesnt pick up just portions of words
boolSpaceAdded = True
StartSearch:
lngFoundCount = 0 'set the found count to zero
arrByteTarget = strdata 'assign the target data to the array
arrByteFind = strValue 'assign the search data to the array
boolFound = False 'set the default found value
lngFindStart = LBound(arrByteFind) 'set the start value
For lngLoopTarget = LBound(arrByteTarget) To UBound(arrByteTarget) Step 1 'start loop through the array byte by byte
  If arrByteFind(lngFindStart) = arrByteTarget(lngLoopTarget) Then 'compare first byte of search string till a match found
    lngStepTemp = lngLoopTarget + 1 'match found so check the rest of the word
    boolFoundTemp = True
    For lngloopStep = (lngFindStart + 1) To UBound(arrByteFind) Step 1
      If lngStepTemp = UBound(arrByteTarget) And lngloopStep < (UBound(arrByteFind)) Then 'if a match is lost before the end of the search string then no match is found
        boolFoundTemp = False
        Exit For
      End If
      If arrByteFind(lngloopStep) <> arrByteTarget(lngStepTemp) Then
        boolFoundTemp = False
        Exit For
      End If
      lngStepTemp = lngStepTemp + 1
    Next lngloopStep
    If boolFoundTemp = True Then 'if there was a match found
      If lngFoundCount > 0 Then 'check to see if this is the first match
        ReDim Preserve arrPos(UBound(arrPos) + 1) 'add the start position to the array
      Else
        ReDim arrPos(0) 'if this is the first match initialise the array
      End If
      If boolSpaceAdded = False Then
        arrPos(UBound(arrPos)) = (lngLoopTarget / 2) 'if no padding was added calculate position
      Else
        arrPos(UBound(arrPos)) = (lngLoopTarget / 2) + 1 'padding added calculate position
      End If
      lngFoundCount = lngFoundCount + 1 'increment count
      boolFound = True 'set match found to true
    End If
  End If
Next lngLoopTarget
If boolFound = True Then 'there was a match found
  If boolRemove = True Then 'check if it is to be removed
    If boolSpaceAdded = True Then 'check the padding
      lngStrValBytCount = ((Len(strValue) - 1) * 2)
    Else
      lngStrValBytCount = (Len(strValue) * 2)
    End If
    For lngLoopTarget = 0 To (lngFoundCount - 1) 'Fill the left hand side temp array with data to the left of a match
      If lngLoopTarget > 0 Then
        lngTempBound = ((((arrPos(lngLoopTarget) * 2) - 2)) - (lngStrValBytCount * lngLoopTarget)) 'caclulate the position in the array of the match
      Else
        lngTempBound = ((arrPos(lngLoopTarget) * 2) - 2)
      End If
      For lngStepTemp = LBound(arrByteTarget) To lngTempBound Step 1 'fill the array
        If lngStepTemp = LBound(arrByteTarget) Then
          ReDim arrByteTempLeft(0)
        Else
          ReDim Preserve arrByteTempLeft(UBound(arrByteTempLeft) + 1)
        End If
        arrByteTempLeft(lngStepTemp) = arrByteTarget(lngStepTemp)
      Next lngStepTemp
      If lngLoopTarget > 0 Then 'calculate the start position of the right hand side of the match
        lngTempBound = (((arrPos(lngLoopTarget) * 2) - 2) - (lngStrValBytCount * lngLoopTarget) + lngStrValBytCount)
        Else
        lngTempBound = (((arrPos(lngLoopTarget) * 2) - 2) + lngStrValBytCount)
      End If
      For lngStepTemp = lngTempBound To UBound(arrByteTarget) Step 1 'fill the array
        If lngStepTemp = lngTempBound Then
          ReDim arrByteTempRight(0)
        Else
          ReDim Preserve arrByteTempRight(UBound(arrByteTempRight) + 1)
        End If
        arrByteTempRight(UBound(arrByteTempRight)) = arrByteTarget(lngStepTemp)
      Next lngStepTemp
      arrByteTarget = arrByteTempLeft
      lngStepTemp = UBound(arrByteTarget) 'join the two halves back together now that a match item has been removed
      ReDim Preserve arrByteTarget(((UBound(arrByteTarget)) + (UBound(arrByteTempRight))))
      For lngloopStep = LBound(arrByteTempRight) To UBound(arrByteTempRight)
        arrByteTarget(lngStepTemp) = arrByteTempRight(lngloopStep)
        lngStepTemp = lngStepTemp + 1
      Next lngloopStep
    Next lngLoopTarget 'loop through all matches in array
    strdata = "" 'prepare target string
    For lngloopStep = LBound(arrByteTarget) To UBound(arrByteTarget) Step 1 'fill string
      If arrByteTarget(lngloopStep) > 0 Then
        strdata = strdata & Chr$(arrByteTarget(lngloopStep))
      End If
    Next lngloopStep
  End If
  If boolReplace = True Then 'if replace was selected
    If boolSpaceAdded = True Then 'check padding
      lngStrValBytCount = ((Len(strValue) - 1) * 2)
    Else
      lngStrValBytCount = (Len(strValue) * 2)
    End If
    lngStrRemBytCount = (Len(strReplace) * 2)
    For lngLoopTarget = 0 To (lngFoundCount - 1)
      If lngLoopTarget > 0 Then 'calculate match position
        lngTempBound = (((arrPos(lngLoopTarget) * 2) - 2)) - (lngStrValBytCount * lngLoopTarget)
        lngTempBound = lngTempBound + (lngStrRemBytCount * lngLoopTarget) - 2
      Else
        lngTempBound = ((arrPos(lngLoopTarget) * 2) - 2)
      End If
      For lngStepTemp = LBound(arrByteTarget) To lngTempBound Step 1 'fill left have array
        If lngStepTemp = LBound(arrByteTarget) Then
          ReDim arrByteTempLeft(0)
        Else
          ReDim Preserve arrByteTempLeft(UBound(arrByteTempLeft) + 1)
        End If
        arrByteTempLeft(lngStepTemp) = arrByteTarget(lngStepTemp)
      Next lngStepTemp
      If lngLoopTarget > 0 Then 'calculate right hand postion
        lngTempBound = (((arrPos(lngLoopTarget) * 2) - 2) - (lngStrValBytCount * lngLoopTarget) + lngStrValBytCount)
        lngTempBound = lngTempBound + (lngStrRemBytCount * lngLoopTarget) - 2
        Else
        lngTempBound = (((arrPos(lngLoopTarget) * 2) - 2) + lngStrValBytCount)
      End If
      For lngStepTemp = lngTempBound To UBound(arrByteTarget) Step 1 ' fill right hand side array
        If lngStepTemp = lngTempBound Then
          ReDim arrByteTempRight(0)
        Else
          ReDim Preserve arrByteTempRight(UBound(arrByteTempRight) + 1)
        End If
        arrByteTempRight(UBound(arrByteTempRight)) = arrByteTarget(lngStepTemp)
      Next lngStepTemp
      lngStepTemp = UBound(arrByteTempLeft) 'prepare bounds for inserting replacement string
      ReDim Preserve arrByteTempLeft(((UBound(arrByteTempLeft)) + (UBound(arrByteReplace))))
      For lngloopStep = LBound(arrByteReplace) To UBound(arrByteReplace) 'insert replacement string
        arrByteTempLeft(lngStepTemp) = arrByteReplace(lngloopStep)
        lngStepTemp = lngStepTemp + 1
      Next lngloopStep
      arrByteTarget = arrByteTempLeft
      lngStepTemp = UBound(arrByteTarget)
      ReDim Preserve arrByteTarget(((UBound(arrByteTarget)) + (UBound(arrByteTempRight)))) 'join arrays again
      For lngloopStep = LBound(arrByteTempRight) To UBound(arrByteTempRight)
        arrByteTarget(lngStepTemp) = arrByteTempRight(lngloopStep)
        lngStepTemp = lngStepTemp + 1
      Next lngloopStep
    Next lngLoopTarget
    strdata = "" 'prepare string
    For lngloopStep = LBound(arrByteTarget) To UBound(arrByteTarget) Step 1 'fill string
      If arrByteTarget(lngloopStep) > 0 Then
        strdata = strdata & Chr$(arrByteTarget(lngloopStep))
      End If
    Next lngloopStep
  End If
  findInString = 1 'success
  GoTo ExitFunction
Else
  findInString = 0 'no match found
  GoTo ExitFunction
End If
ErrorHandler:
  findInString = 6 'oops I hope that wasnt my fault
ExitFunction:
'clean up after ourselves
Erase arrByteFind
Erase arrByteTarget
Erase arrByteReplace
Erase arrByteTempRight
Erase arrByteTempLeft
lngLoopTarget = vbNull
lngloopStep = vbNull
lngStepTemp = vbNull
lngFindStart = vbNull
lngStrValBytCount = vbNull
lngTempBound = vbNull
lngStrRemBytCount = vbNull
End Function
```

