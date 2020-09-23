<div align="center">

## Set Icons for any Form or Program


</div>

### Description

With this code you can place any Icon in the title bar of any Window, just by reffering to a .ico file or to the position of the Icon in a DLL.
 
### More Info
 
Handle of the window you want to change the icon of.

You need to have the Window Handle (hWnd) of the window whitch Icon you want to change. This can be done by searching/finding it with the API-call

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

none (that I know of)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[J\. van Gils](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/j-van-gils.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/j-van-gils-set-icons-for-any-form-or-program__1-1912/archive/master.zip)

### API Declarations

```
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_SETICON = &H80
```


### Source Code

```
Public Function SetIcon(FormhWnd As Long)
Dim x, i As Long
  i = ExtractIcon(0, "c:\SomeDll.DLL", 3)
   'In this case you will extract the 3rd icon from SomeDll.DLL. In this
   'way you can extract any icon you want, just by reffering to the icon
   '(number) of the icon you want to extract in the dll. If you want to
   'know the iconnumbers of a dll, you will have to use a recource editor
   '(like Borland Recource Workshop). You can also extract the Icon Handle
   'of a .ico file just by using some code like:
   'i=ExtractIcon(0,"c:\SomeIconFile.ico",0)
   'where SomeIconFile is the name of the icon you want to use.
   'Now finally set the icon in the title bar of the window
  x = DefWindowProc(FormhWnd, WM_SETICON, &H1, i)
End Function
```

