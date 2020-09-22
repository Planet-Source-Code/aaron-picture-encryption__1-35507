<div align="center">

## Picture Encryption


</div>

### Description

This will encrypt common text into a picture, 3 characters per pixel. You can save/open the picture and text. Works Great! Any suggestions would be greatly appreciated.
 
### More Info
 
Conflicting color palettes/display modes may cause the program to decrypt incorrectly on some computers - make sure your vid card is set to at least 24bit!


<span>             |<span>
---                |---
**Submitted On**   |2002-06-05 13:54:34
**By**             |[Aaron](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aaron.md)
**Level**          |Intermediate
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Picture\_En90693652002\.zip](https://github.com/Planet-Source-Code/aaron-picture-encryption__1-35507/archive/master.zip)

### API Declarations

```
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
```





