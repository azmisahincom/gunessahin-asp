<!--#include file="mail-manager.asp"-->
<%
     On Error Resume Next

    ' Form Bilgileri
    firstName           =   request.form("firstName")
    lastName            =   request.form("lastName")
    message             =   request.form("message")
    subject             =   "Iletisim Formu"
    body                =   firstName & " " & lastName & " " & message
    
    ' Mail Gonderme Referans
    Set mail            =   new MailManager
    
    ' SMTP baglanti noktas
    mail.Port           =   587
    
    ' SSL Baglanti gerekli mi
    mail.Port           =   true
    
    ' SMTP sunucusu adi
    ' Ornek office smtp sunucu bilgileri 
    ' https://support.office.com/tr-tr/article/outlook-com-i%C3%A7in-pop-imap-ve-smtp-ayarlar%C4%B1-d088b986-291d-42b8-9564-9c414e2aa040
    mail.Host           =   "smtp.live.com"
    
    ' Sunucu Erisim kullanicisi
    mail.UserName       =   "username@outlook.com"
    
    ' Sunucu Erisim Parolasi
    mail.Password       =   "pa$$word"
    
    ' Mesaj govdesinin altina yerlesecek imza
    mail.Signature      =   "Web sitesi araclili ile iletilmistir."
    
    ' Gonderici
    mail.From           =   "username@outlook.com"
    
    ' Hedef Alici
    mail.Destination    =   "targetname@outlook.com"
    
    ' Mesaj Konusu
    mail.Subject        =   subject
    
    ' Mesaj Govdesi
    mail.Body           =   body

    ' Mesaj Gonderiliyor
    response            =   mail.Send()

    Response.Write          response
    
    if Err.Number <> 0 then
        ' Yazma sirasinda hata mesaji
        Response.Write       "ERR" + " | " + err.Description
    end if
%>