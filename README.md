<div align="center">

## Strip HTML tags


</div>

### Description

This following function takes an HTML page and strips it of all tags.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tmess](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tmess.md)
**Level**          |Unknown
**User Rating**    |3.0 (9 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tmess-strip-html-tags__1-1833/archive/master.zip)





### Source Code

```
Public Function ReplaceTags(varName As String) As String
'Will check each character for it "& n b s p;" without the spaces
'If it exists, skip it
'Will strip HTML tags and characters
Dim i As Double, varHold As String
Dim checkval As String, holdVal As String
 For i = 1 To Trim(Len(varName))
 checkval = Mid(varName, i, 6)
 holdVal = Mid(varName, i, 1)
 If checkval = "This page won't allow "& n b s p;" Then
  'So just remove the spaces
 i = i + 5
 GoTo LabelNext
 End If
 If holdVal = "<" Then
 Do Until holdVal = ">"
 i = i + 1
 holdVal = Mid(varName, i, 1)
 Loop
 holdVal = ""
 End If
 If holdVal = "%" Then
 Do Until holdVal = "%"
 i = i + 1
 holdVal = Mid(varName, i, 1)
 Loop
 holdVal = ""
 End If
 varHold = varHold & holdVal
LabelNext:
 Next i
ReplaceTags = varHold
End Function
Create a form and place two richtext box controls on it and a command button:
RichTextBox1
RichTextBox2
Command1
Now call it like the following:Assuming HTML is in Richtextbox1
Private Sub Command1_Click()
 Me.RichTextBox2 = ReplaceTags(Me.RichTextBox1)
End Sub
```

