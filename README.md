<div align="center">

## add a horizontal scroll bar to a listbox or combo


</div>

### Description

add a horizontal scroll bar to a listbox or combo box
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Qaid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-qaid.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-qaid-add-a-horizontal-scroll-bar-to-a-listbox-or-combo__1-179/archive/master.zip)

### API Declarations

```
#If Win16 Then
  Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#Else
  Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If
```


### Source Code

```
Const WM_USER = 1024
Const LB_SETHORIZONTALEXTENT = (WM_USER + 21)
Dim nRet As Long
Dim nNewWidth As Integer
nNewWidth = list1.Width + 100 'new width in pixels
nRet = SendMessage(list1.hwnd, LB_SETHORIZONTALEXTENT, nNewWidth, ByVal 0&)
```

