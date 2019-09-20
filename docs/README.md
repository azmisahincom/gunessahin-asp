# ![Logo](media/favicon.png)

## Simple File Write

```ASP
<%  
    ' Simple File Write
    ' Copyright (c) 2019 Güneş ŞAHİN
    ' https://github.com/gunessahin

    Dim fileStream, file
    Set fileStream  =   Server.CreateObject("Scripting.FileSystemObject")
    defaultPath     =   "/media"
    fileNameString  =   "deneme.txt"    
    fileName        =   Server.MapPath(defaultPath + "/" + fileNameString)   

    On Error Resume Next
            
    fileContent     =   "HELLO WORLD ! " & vbCrlf & Now()

    Set file        =   fileStream.CreateTextFile(fileName, true)

    file.WriteLine      fileContent

    if Err.Number <> 0 then
        ' Yazma sırasında hata mesajı
       Response.Write       "ERR" + " | " + err.Description
    end if

    file.close
    Set file        =   Nothing
    Set fileStream  =   Nothing

%>
```