<div align="center">

## Detect if cookies are enabled


</div>

### Description

When you work with cookies, you always run into paranoid types who disable their cookies. If you don't detect and deal with them, your code may not work. This code detects the user's cookie settings using the ASP Session object. Unlike some other implementations, it requires only one script page.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Beginner
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-detect-if-cookies-are-enabled__4-6092/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<%option explicit
if (Request("lngSessionId")="") then
	'first time through
	'redirect back to ourselves
	Response.Redirect "CookieCheck.asp?lngSessionId=" & Session.SessionID
else
	'second time through
dim lngSessionId
	lngSessionId = Request("lngSessionId")
	If lngSessionId = Session.SessionID then
		Response.Write "User has cookies turned ON."
	Else
		Response.Write "User is paranoid and has cookies turned OFF."
	End If
end if
%>
```

