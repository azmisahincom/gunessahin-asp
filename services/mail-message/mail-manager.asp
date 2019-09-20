<%
    ' Mail Manager
    ' Copyright (c) 2019 Güneş ŞAHİN
    ' https://github.com/gunessahin
    '================================================================================
    
    ' Mail Manager Sınıfı
    class MailManager
        
        ' SINIF TANIMLARI
        ' ================================================================================
    
        public Port                     '   Servis Portu
        public EnableSsl                '   SSL Bağlantı gerekli mi
        
        public Host                     '   Smtp Sunucusu
        public UserName                 '   Sunucu Erişim kullanıcısı
        public Password                 '   Sunucu Erişim Parolası
        public Domain                   '   
        public Signature                '   Mesaj gövdesinin altına yerleşecek imza
        
        public From                     '   Gönderici
        public Destination              '   Hedef Alıcı
        public Subject                  '   Mesaj Konusu
        public Body                     '   Mesaj Gövdesi

        ' Dönüş bilgisi
        public Result                   '   ServicesResult

            ' Sınıf İlk Yapılandırması
        public sub Class_Initialize()

            ' Varsayılan Yeni Servis Dönüş Modeli
            Set Result                  =   "Gonderiliyor"

        end sub

        public function Send()
    
            Dim oMessage
            Dim oConfig
    
            ' Referans : https://msdn.microsoft.com/en-us/library/ms992546(v=exchg.65).aspx
            Const cdoSendUsingPort      =   2
        
            ' Collaboration Data Objects
            ' Refarans : https://msdn.microsoft.com/en-us/library/cc161087(v=exchg.65).aspx
            Set oMessage                =   Server.CreateObject("CDO.Message")

            ' Referans : https://msdn.microsoft.com/en-us/library/ms526318(v=exchg.10).aspx
            Set oConfig                 =   Server.CreateObject ("CDO.Configuration")
    
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")                 =   Host
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")             =   Port
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")                  =   cdoSendUsingPort
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")      =   10
    
            ' Authenticated method
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")           =   1
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")                 =   EnableSsl
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")               =   UserName
            oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")               =   Password
            oConfig.Fields.Update
    
            Set oMessage.Configuration                                                                  =   oConfig
    
            ' From Aynı Zamanda Giriş Yapan Mail Adresi
            oMessage.From                                                                               =   From
    
            ' Hedef Alıcı
            oMessage.To                                                                                 =   Destination
    
            ' Mesaj Konusu
            oMessage.Subject                                                                            =   Subject
    
            ' Body Daha önceden Html olarak yapılandırılmış verileri iletir.
            oMessage.HtmlBody                                                                           =   Body
    
            ' Gönderim Sırasında Hata Durumları
            On Error Resume Next
    
            ' Mesajı Gönder
            oMessage.Send

            ' Gönderme Sırasında Bir Hata Oluştumu    
            If Err <> 0 Then

                ' Mesaj Gönderilemedi                
                Result                                                                                  =   Err.Description

            Else
                ' Mesaj Gönderildi
                Result                                                                                  =   "OK"
    
            End If
    
            Set oMessage                                                                                =   Nothing
            Set oConfig                                                                                 =   Nothing
    
            ' Fonksiyon Dönüşü
            Send                                                                                        =   Result

        end function    

    ' Sınıf Sonu
    end class
%>