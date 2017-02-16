<%
Set pc = CreateObject("Wscript.Network") 
    'response.write "nome computador: " & pc.ComputerName & "<br>" 
	Nomecomputador = pc.ComputerName 

Dim mail, body

body = "Name:  asdfasf afsda fdasfasdf afdasfadsf saf dafdasf dasf"

Set objCDOConf = Server.CreateObject ("CDO.Configuration")
' ** SET AND UPDATE FIELDS PROPERTIES **
With objCDOConf
    ' ** OUT GOING SMTP SERVER **
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Your SMTP.com or SMTP IP"
    ' ** SMTP PORT **
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    ' ** CDO PORT **
    .Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    ' ** TIMEOUT **
    .Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    .Fields.Update 
End With

Set objMail = Server.CreateObject("CDO.Message")

' ** UPDATE THE CDOSYS CONFIGURATION **
Set objMail.Configuration = objCDOConf

' ** Set mail = Server.CreateObject("CDO.Message") **
objMail.To = "destination@email.com"
objMail.From = "teste@teste.com"
objMail.Subject = "Tete de envio de email do SERVIDOR: " & Nomecomputador
objMail.TextBody = "Teste de envio de email do Servidor: " &Nomecomputador & " Horario: " & now
objMail.Send()

Response.Write  "Teste de envio de email do Servidor: " & Nomecomputador & "<br> Horario: " & now

Set objMail = nothing
Set body = nothing
Set objCDOConf = Nothing

%>
