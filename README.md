<div align="center">

## Execute a file Correctly


</div>

### Description

It correctly executes a program
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[AdamRC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adamrc.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adamrc-execute-a-file-correctly__1-13645/archive/master.zip)





### Source Code

```
DirPath = [Path of file]
  On Error GoTo err:
  X% = Shell(DirPath, 1): NoFreeze% = DoEvents(): Exit Sub
  Exit Sub
err:
  If err.Number = 6 Then Exit Sub
  MsgBox "Please make sure you have the correct path and then try again."
```

