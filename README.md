<div align="center">

## EBrowseF\.zip


</div>

### Description

Enhancing the SHBrowseForFolder API Function<br>

----

<br>

Browse for a Folder using SHBrowseForFolder API function with a callback function BrowseCallbackProc. This Extends the functionality that was given in the MSDN Knowledge Base article Q179497 "HOWTO: Select a Directory Without the Common Dialog Control".

After reading the MSDN knowledge base article Q179378 "HOWTO: Browse for Folders from the Current Directory", I was able to figure out how to add a callback function that sets the starting directory and displays the currently selected path in the "Browse For Folder" dialog.
 
### More Info
 
Public Function BrowseForFolder(owner As Form, Title As String, StartDir As String) As String

owner = referance to the calling from (can be 0)

Title = Intructions to appear in the "Browse for Folder" dialog.

StartDir = The directory that is displayed in the tree view when the dialog appears.

How to call API functions and how to use the AddressOf operator to get a function pointer and store it in a variable.

If user doesn't click cancel then it returns a string with the directory the users selected.


<span>             |<span>
---                |---
**Submitted On**   |2000-02-20 15:03:58
**By**             |[Stephen Fonnesbeck](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephen-fonnesbeck.md)
**Level**          |Advanced
**User Rating**    |5.0 (60 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD35302212000\.zip](https://github.com/Planet-Source-Code/stephen-fonnesbeck-ebrowsef-zip__1-6175/archive/master.zip)








