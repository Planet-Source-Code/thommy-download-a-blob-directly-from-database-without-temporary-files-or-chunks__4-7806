<div align="center">

## Download a BLOB directly from Database without temporary files or chunks


</div>

### Description

Download a binary large object (BLOB) from the database without temporary files or chunks.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thommy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thommy.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thommy-download-a-blob-directly-from-database-without-temporary-files-or-chunks__4-7806/archive/master.zip)





### Source Code

```
'Declarations
Dim objStream 'As ADODB.Stream
Dim objBLOB  'As Object
Dim strBLOB  'As String
Dim lngBLOB  'As Long
'Create objects
Set objStream = Server.CreateObject("ADODB.Stream")
'Get the object, the object size and the object name from the database
'Set objBLOB = ... 'Get the BLOB out of the database
'strBLOB = ...   'Filename to be saved to
'lngBLOB = ...   'Length of the BLOB
'Assign object to stream
With objStream
	.Type = 1
	.Open
	.Write objBLOB
	.Position=0
End With
'Start forcing download
With Response
	.AddHeader "Content-Disposition", "attachment; filename=" & strBLOB
	.AddHeader "Content-Length", lngBLOB
	.ContentType = "binary/octet-stream"
	.BinaryWrite objStream.Read
	.Flush
End With
'Close all objects
objStream.Close
'Clean up
Set objStream = Nothing
Set objBLOB = Nothing
```

