Function sendMail(aTo, aSub, aText)
On Error resume Next
	Dim ObjSendMail
	Set ObjSendMail = CreateObject("CDO.Message") 
    ObjSendMail.BodyPart.Charset = "UTF-8" 
	'This section provides the configuration information for the remote SMTP server.
		 
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="mail.iranmining.com"
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		 
	' If your server requires outgoing authentication uncomment the lines below and use a valid email address and password.
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="noreply@iranmining.com"
	ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="sharp792c"
		 
	ObjSendMail.Configuration.Fields.Update
		 
	'End remote SMTP server configuration section==
		 
	ObjSendMail.To = aTo
	ObjSendMail.Subject = aSub
	ObjSendMail.From = """Iran Mining""<noreply@iranmining.com>"
		 
	' we are sending a text email.. simply switch the comments around to send an html email instead
	ObjSendMail.HTMLBody = aText
	'ObjSendMail.TextBody = aText
	'ObjSendMail.Importance = 1
	ObjSendMail.Send
		 
	'Set ObjSendMail = Nothing 
	If Err <> 0 Then
		Response.Write "Error encountered: " & Err.Description
	 End If
	 Set ObjSendMail = Nothing
	 sendMail=Err.Number

	 'Dim ObjSendMail
	 'Set ObjSendMail = Server.CreateObject("Persits.MailSender")
	 
	 'ObjSendMail.Host = "mail.iranmining.com"
	 'ObjSendMail.From = "noreply@iranmining.com"
	 'ObjSendMail.FromName = "noreply"
	 'ObjSendMail.AddAddress aTo
	 'ObjSendMail.Subject = ObjSendMail.EncodeHeader(aSub, "utf-8")
	 'ObjSendMail.Body = aText
	 'ObjSendMail.isHTML = True
	 'ObjSendMail.CharSet = "UTF-8"

	 'ObjSendMail.Username = "noreply"
	 'ObjSendMail.Password = "sharp792c"
	 
	 'On Error Resume Next
	 'ObjSendMail.ContentTransferEncoding = "Quoted-Printable"
	 'ObjSendMail.Send
	 'If Err <> 0 Then
		'Response.Write "Error encountered: " & Err.Description
	 'End If
	 'Set ObjSendMail = Nothing
	 'sendMail=Err.Number
End Function 

sendMail ("my.samimi@hmail.com","test","Hi, This is test!")