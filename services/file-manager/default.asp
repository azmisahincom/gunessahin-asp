<!--#include file="file-manager.asp"-->
<%
    On Error Resume Next

    Dim fileExplorer
    Set fileExplorer                    =   new FileManager

    fileExplorer.Browse()
%>