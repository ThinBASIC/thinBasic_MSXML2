# thinBasic_MSXML2
thinBasic_MSXML2 enables thinBasic programmer to use selected MS XML 2.0 objects.

## For user
### Implemented classes
ServerXMLHTTPRequest module class implements MS [IServerXMLHTTPRequest](https://msdn.microsoft.com/en-us/library/ms754586(v=vs.85).aspx) interface.

DOMDocument module class implements MS [IXMLDOMDocument](https://msdn.microsoft.com/en-us/library/ms756987(v=vs.85).aspx) interface.

## For developer
### Language
PowerBASIC for Windows, v10.04
* Default compiler Win32 API headers used
* COM interfaces generated by *PowerBASIC COM Browser*

### Dependencies
Clone [module_core](https://github.com/ThinBASIC/module_core) in a way it is placed in the same root directory as this project.

### Script example using this module
...
  Uses "MSXML2"
  Uses "Console"

  '---Reference for Fake JSon: http://jsonplaceholder.typicode.com/
  
  Dim oHTTP As new ServerXMLHTTPRequest
  printl "IsObject:", oHTTP.IsObject
  printl "IsNothing:", oHTTP.IsNothing

  oHTTP.SetTimeOuts(60000, 60000, 60000, 60000)
  '------------------------------------------------------------
  ' Users
  '------------------------------------------------------------
  printl "---Users---" in %CColor_fYellow
  oHTTP.Open("GET", "http://jsonplaceholder.typicode.com/users", %FALSE)
  oHTTP.Send
  PrintL "Status:", oHTTP.Status, "(" & oHTTP.Statustext & ")"
  PrintL oHTTP.ResponseText

  '------------------------------------------------------------
  ' Single Post
  '------------------------------------------------------------
  printl "---Single Post---" in %CColor_fYellow
  oHTTP.Open("GET", "http://jsonplaceholder.typicode.com/posts/1", %FALSE)
  oHTTP.Send
  PrintL "Status:", oHTTP.Status, "(" & oHTTP.Statustext & ")"
  PrintL oHTTP.ResponseText

  '------------------------------------------------------------
  ' All Posts
  '------------------------------------------------------------
  printl "---All Posts---" in %CColor_fYellow
  oHTTP.Open("GET", "http://jsonplaceholder.typicode.com/posts", %FALSE)
  oHTTP.Send
  PrintL "Status:", oHTTP.Status, "(" & oHTTP.Statustext & ")"
  PrintL oHTTP.ResponseText
  
  PrintL
  WaitKey
...

