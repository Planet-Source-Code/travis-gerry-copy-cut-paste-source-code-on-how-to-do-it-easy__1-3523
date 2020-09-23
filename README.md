<div align="center">

## COPY, CUT, PASTE Source code on how to do it EASY


</div>

### Description

This Code shows you how to copy, cut, and paste with only a few lines this is very easy to do. This is the CODE and not just some garbage like those other copy, paste, and cut. AND IT WORKS!!!
 
### More Info
 
You need to know the basics of Visual Basic Programing, and thats it


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Travis Gerry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/travis-gerry.md)
**Level**          |Unknown
**User Rating**    |4.5 (54 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/travis-gerry-copy-cut-paste-source-code-on-how-to-do-it-easy__1-3523/archive/master.zip)





### Source Code

```
' You need a menu with 3 options, cut, copy, paste. Make sure to name them
' mnucut, mnucopy, mnupaste. Or just make 3 buttons and change the code a bit.
' You need one text box, defualt name text1. And thats it.
Private Sub mnucopy_Click()
If Text1.SelText = "" Then
  Exit Sub
Else
  Clipboard.Clear
  Clipboard.SetText Text1.SelText
End If
End Sub
Private Sub mnucut_Click()
If Text1.SelText = "" Then
  Exit Sub
Else
  Clipboard.Clear
  Clipboard.SetText Text1.SelText
  Text1.SelText = ""
End If
End Sub
Private Sub mnupaste_Click()
Text1.SelText = Clipboard.GetText
End Sub
```

