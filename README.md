<div align="center">

## Change selected printer


</div>

### Description

Temporarily change the currently selected (default) printer within your VB program
 
### More Info
 
This doesn't save the change to the registry - which needs to be done for a permanent change of default printer.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Beginner
**User Rating**    |4.1 (45 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-change-selected-printer__1-29428/archive/master.zip)





### Source Code

```
Public Function SetDefaultPrinter(ByVal DeviceName As String) As Boolean
Dim prThis As Printer
If Printers.Count > 0 Then
  '\\ Iterate through all the installed printers
  For Each prThis In Printers
    '\\ If the desired one is found
    If prThis.DeviceName = DeviceName Then
      Set Printer = prThis
      SetDefaultPrinter = True
      '\\ Stop searching
      Exit For
    End If
  Next prThis
End If
End Function
```

