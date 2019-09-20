<%
    ' File Manager
    ' Copyright (c) 2019 Güneş ŞAHİN
    ' https://github.com/gunessahin
    '================================================================================
    
    ' File Manager Sınıfı
    class FileManager
    
        ' SINIF TANIMLARI
        ' ================================================================================
        public FileStream
    
        public Overwrite
        public Unicode
        public DefaultPath
        public FileNameString
        public ServerPath
        public FileName
        public FileContent
    
        ' Sınıf İlk Yapılandırması
        public sub Class_Initialize()

            ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
            Set FileStream  =   Server.CreateObject("Scripting.FileSystemObject")

            ' Varsayılan Değerler
            Overwrite       =   true
            Unicode         =   false
            DefaultPath     =   "/media"
            FileNameString  =   "test.txt"
            ServerPath      =   Server.MapPath(DefaultPath + "/" + FileNameString)
            FileName        =   ServerPath
            FileContent     =   "TEST"

        end sub

        ' Yazılacak klasörü tanımlar.
        public function setPath(path)
            FilePath        =   path
        end function

        ' Yazılacak dosyası tanımlar, üzerine yazmayı etkinleştirir.
        public function setFile(file, over)

            FileNameString  =   file
            ServerPath      =   Server.MapPath(DefaultPath + "/" + FileNameString)
            FileName        =   ServerPath
            Overwrite       =   over
        end function

        ' String içeriği dosyaya yazar.
        public function WriteString(content)
                
            On Error Resume Next
            
            FileContent     =   content

            Set file            =   FileStream.CreateTextFile(FileName, Overwrite, Unicode)

            file.WriteLine          FileContent
    
            if Err.Number <> 0 then
                ' Yazma sırasında hata mesajı
               Response.Write       "ERR" + " | " + err.Description
            end if

            file.Close()
            Set file            =   Nothing

            ' Yazma sonrası mesajı
            Response.Write          "OK" + " | " + FileName
    
        end function

        ' Data içeriği dosyaya yazar.
        public function WriteData(data)
                
            On Error Resume Next
          
            Set file            =   FileStream.CreateTextFile(FileName, Overwrite, Unicode)

            file.write              data
    
            if Err.Number <> 0 then
                ' Yazma sırasında hata mesajı
               Response.Write       "ERR" + " | " + err.Description
            end if

            file.Close()
            Set file            =   Nothing

            ' Yazma sonrası mesajı
            Response.Write          "OK" + " | " + FileName
    
        end function

        ' Klasör içerisinde dosyaları listeler
        public function Browse()
             
            dirPath = server.MapPath(DefaultPath)
            
            Set folder = FileStream.GetFolder(dirPath)
       
            Set files = folder.Files
            
            Response.Write "<ul>"
            For Each item In files
                serverAddress = Request.ServerVariables("server_name") & ":" & Request.ServerVariables("server_port")
                serverFilePath = "//" & serverAddress & DefaultPath & "/" & item.Name
                Response.Write "<li><a href='" & serverFilePath & "'>" & item.Name & "</a>" & "</li>"
            Next      
            Response.Write "</ul>"           

        end function

        ' Upload file
        public function Upload(Request)
    
    		totalBytes      =   Request.TotalBytes
            binaryData      =   Request.BinaryRead (totalBytes)
    
            'Set fso = CreateObject("Scripting.FileSystemObject")
            'Set f = fso.OpenTextFile(SavePath & "\" & FileName, ForWriting, True)
			'f.Write strFileData
			'Set f = nothing
			'Set fso = nothing

            'setFile             "deneme.txt" , false

            'WriteData           binaryData
        
            Response.Write("File " & totalBytes & " byte ")

        end function

    ' Sınıf Sonu
    end class
%>