<!--#include file="file-manager.asp"-->
<%
    On Error Resume Next

    fileName                            =   request.Form("fileName")
    fileContent                         =   request.Form("fileContent")

    Dim fileExplorer
    Set fileExplorer                     =  new FileManager

    ' Varsayılan yazma dosya deposu klasörü
    fileExplorer.setPath                    "/media"

    ' Yazılacak dosya adı, Dosya mevcut ise üzerine yazılsın mı?
    fileExplorer.setFile                    fileName, true

    ' Yazılacak dosya içeriği
    fileExplorer.WriteString                fileContent
    
    if Err.Number <> 0 then
        Response.Write                      Err.Description
    end if
%>