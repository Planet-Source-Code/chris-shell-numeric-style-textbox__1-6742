<div align="center">

## Numeric Style Textbox


</div>

### Description

From CShellVB http://www.cshellvb.com

Change the style of a normal textbox so that it will only accept numbers. Better then evaluating keypress events and very fast. This call works on the fly and it is wrapped into a function for you.
 
### More Info
 
NumberText as Text Box and Flag as Boolean

This changes the style of the textbox for the life of the form. Even if the form or textbox is refreshed the style will remain in effect!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Shell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-shell.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-shell-numeric-style-textbox__1-6742/archive/master.zip)

### API Declarations

```
'From CShellVB http://www.cshellvb.com
'Below is used for making a TextBox control accept
'only numbers. Very cool because it changes the style
'and does not require any code on the TextBox's events
'Place declares in Module
Declare Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
'Used for forcing only numbers in a textbox
Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&
```


### Source Code

```
'This function changes the style based on the flag
Public Sub SetNumber(NumberText As TextBox, Flag As Boolean)
Dim curstyle As Long
Dim newstyle As Long
'This Function uses 2 API functions to set the style of
'a textbox so it will only accept numbers CShell
curstyle = GetWindowLong(NumberText.hwnd, GWL_STYLE)
If Flag Then
  curstyle = curstyle Or ES_NUMBER
Else
  curstyle = curstyle And (Not ES_NUMBER)
End If
newstyle = SetWindowLong(NumberText.hwnd, GWL_STYLE, curstyle)
NumberText.Refresh
End Sub
```

