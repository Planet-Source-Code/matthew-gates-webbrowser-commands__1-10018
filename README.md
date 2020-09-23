<div align="center">

## WebBrowser Commands


</div>

### Description

This is a list of some basic WebBrowser Commands and a few advanced ones.
 
### More Info
 
Shdocvw.dll and Comdlg32.dll required.

I didn't bother to make this a Tutorial because I didn't think it belonged in that category. Most of the codes are commented on what it does.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Gates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-gates.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-gates-webbrowser-commands__1-10018/archive/master.zip)





### Source Code

```
'Go to the previous webpage
WebBrowser1.GoBack
'Go to the present webpage
WebBrowser1.GoForward
'Go to the default IE home
WebBrowser1.GoHome
'Go to the default search page
WebBrowser1.GoSearch
'Refresh the current webpage
WebBrowser1.Refresh
'Navigate to a webpage
WebBrowser1.Navigate "www.YourSite.com"
'Printing:
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
'Saving a webpage:
WebBrowser1.ExecWB OLECMDID_SAVEAS,OLECMDEXECOPT_PROMPTUSER
'Open a webpage file on located on your computer:
On Error GoTo fileOpenErr
CommonDialog1.CancelError = True
CommonDialog1.flags = &H4& Or &H100& Or cdlOFNPathMustExist Or cdlOFNFileMustExist
CommonDialog1.DialogTitle = "Select File To Open"
CommonDialog1.Filter = "HTM (*.htm)|*.htm|Txt Files (*.txt)|*.txt|Jpg Files (*.jpg)|*.jpg|Gif Files (*.gif)|*.gif|All Files (*.*)|*.*"
CommonDialog1.ShowOpen
webbrowser1.Navigate CommonDialog1.FileName
fileOpenErr:
Exit Sub
'Open a new web browser using your program and not IE:
Dim NewBrowser as Form1
NewBrowser.Show
NewBrowser.Caption = "MyBrowser"
'Load IE Preferences:
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
'An advanced Stop:
If webbrowser1.Busy Then
webbrowser1.Stop
webbrowser1.GoBack
End If
```

