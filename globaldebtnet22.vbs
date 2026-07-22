DSN_DBT = "DSN=debtnetplusiii;UID=sa;PWD=;"

' Config local con credenciales (mailsecrets.vbs no se sube a git)
Set fsoGlobals = CreateObject("Scripting.FileSystemObject")
SecretsFile = fsoGlobals.BuildPath(fsoGlobals.GetParentFolderName(WScript.ScriptFullName), "mailsecrets.vbs")
if fsoGlobals.FileExists(SecretsFile) then
	Set fSecretsGlobals = fsoGlobals.OpenTextFile(SecretsFile, 1)
	ExecuteGlobal fSecretsGlobals.ReadAll()
	fSecretsGlobals.Close
end if

EMAIL_ADM = "cloyola@debtcontrol.cl"

EMAIL_SUP = "raul.gonzalez@debtcontrol.cl"
EMAIL_CALSER = "servicioalcliente@debtcontrol.cl"

MAIL_SERVER = "mail.debtcontrol.cl"
'MAIL_SERVER = "localhost"
'DIR_OUT = "E:\Comun\Ejecutables\salida\"
DIR_OUT = "E:\Comun\Ejecutables\salida\"

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

Sub Showerror()
	Wscript.Echo "Error Number -> " & Err.Number
	Wscript.Echo "Error Source -> " & Err.Source
	Wscript.Echo "Error Desc   -> " & Err.Description
	Err.Clear
End Sub

Sub SendMail(myto, myfrom, mysubject, mycontent, myarchivo)
	'objeto para enviar email
	sch = "http://schemas.microsoft.com/cdo/configuration/"
	Set cdoConfig = CreateObject("CDO.Configuration") 
	With cdoConfig.Fields 
		.Item(sch & "sendusing") = 2 ' cdoSendUsingPort 
		.Item(sch & "smtpserver") = MAIL_SERVER
		.update 
	End With
	Set myMail=CreateObject("CDO.Message")
	myMail.Configuration = cdoConfig 
	myMail.Subject= mysubject
	myMail.From= myfrom
	myMail.To= myto
	myMail.TextBody= mycontent
	mymail.AddAttachment DIR_OUT & myarchivo
	myMail.Send
	set myMail=nothing
	Set cdoConfig=nothing
	If Err.Number <> 0 then
		Call ShowError()
	End If
End Sub
'1 Raul 2 Julian 3 Lorena
Sub SendMailOutlook(myto, myfrom, mysubject, mycontent, myarchivo,idproceso, total)
 Dim Outlook 'As New Outlook.Application
 Dim rs3
 
 Set Outlook = CreateObject("Outlook.Application")
 ssql = "select usuarios.smail from usuarios, usuario_mail where usuarios.slogin = usuario_mail.slogin and idproceso = " & idproceso & " and total = '" & total & "'"
 Set rs3 = cnnBase.Execute(ssql) 

  'Create e new message
  Dim Message 'As Outlook.MailItem
  Set Message = Outlook.CreateItem(0)
  if myfrom = "" then
        myfrom = EMAIL_ADM
  end if

  With Message
    	.Subject = mysubject
    	.Body = mycontent

	if myto <> "" then
	  	.Recipients.Add (myto)
	else
		 .Recipients.Add (EMAIL_ADM)
	end if
	do while not rs3.eof
		if (rs3(0) <> "") then
			.recipients.add (rs3(0))
		end if
		rs3.movenext
	loop
	rs3.close
	if myarchivo <> "" then
	    	.Attachments.Add(DIR_OUT & myarchivo).Displayname = "Resultado"
	end if
    	.Send
  end with
end sub

sub EnviaMailNew ()
	dim rs 
	dim rs1
	dim rs2
	dim rs3
	Dim fso1, f1
	dim bodytext
	dim Nomusaurio 
	dim fonousuario
	dim clienteusuario
	Dim strcc
	dim sql
	Dim nrut
	Dim sdv

       
       On Error resume next

	sql = "select ltrim(rtrim(de)) ,para ,cc,ltrim(rtrim(referencia)) ,texto, idmail, isnull(adjunto,'') , archbody, nkey_mail, ltrim(rtrim(tablaorigen)), idorigen , isnull(iddestinos,0),nkey_cliente , enviado, isnull(html,'N') from mail where enviado = 'X' " 
	sql = sql & " union all "
	sql = sql & " select ltrim(rtrim(de)) ,ltrim(rtrim(para)) ,cc, ltrim(rtrim(referencia)),ltrim(rtrim(texto)), idmail, ltrim(rtrim(isnull(adjunto,''))), '', nkey_mail, ltrim(rtrim(tablaorigen)), idorigen , isnull(iddestinos,0),nkey_cliente,   enviado , isnull(html,'N') from mail where enviado = 'N' "

	Set rs = cnnBase.Execute(sql)
	strcc = ""
	do while not rs.eof
		Nomusaurio =  " "
		fonousuario = " "
		clienteusuario = " "
		sdv = " "
		nrut = " "
		if rs(9) = "GestionDeudor" then
			Set rs1 = cnnBase.Execute("select cliente.holding, usuarios.snombre, analista.anexo, cliente.nrutHolding, cliente.sdigitoHolding  from  cliente, gestiondeudor, usuarios, analista where analista.sloginanalista = usuarios.slogin and cliente.nkey_cliente = gestiondeudor.nkey_cliente and gestiondeudor.nkey_gestiondeudor = " & rs(10) & " and gestiondeudor.slogin  = usuarios.slogin and analista.activo = 'S'  and usuarios.activo = 'S'   ")
			if not rs1.eof then
				Nomusaurio =  rs1(1)
				fonousuario = "Fono (56-2) 2599-" & rs1(2)
				clienteusuario = rs1(0)
				nrut =  rs1(3)
				sdv =  " - " & rs1(4)
			end if
		end  if


		Set objMessage = CreateObject("CDO.Message") 
		objMessage.Subject = rs(3)
		maildesde = rs(0)
		if rs(0) <> "" then
			Set rs1 = cnnBase.Execute("select ltrim(rtrim(mailuserori)) as 'mailuserori', ltrim(rtrim(mailpassori )) as 'mailpassori' from cliente where mailorigen = '" & rs(0) & "'" )
			if rs1.eof then
                                Set rs1 = cnnBase.Execute("select ltrim(rtrim(mailuserori)) as 'mailuserori', ltrim(rtrim(mailpassori )) as 'mailpassori' from cliente where mailorigen = 'servicioalcliente@debtcontrol.cl'" )
                                maildesde = "servicioalcliente@debtcontrol.cl" 
			end if
			objMessage.From = maildesde
			objMessage.To = rs(1)
			If rs(13) = "N" Then
				Set rs3 = cnnBase.Execute("select ltrim(rtrim(isnull(usuarios.smail,'" & EMAIL_ADM & "')))  from usuarios, usuario_mail where usuarios.slogin = usuario_mail.slogin and idproceso = " & rs(11) & " and usuario_mail.nkey_cliente in (0, " & rs(12) & ")")
				do while not rs3.eof
					If strcc <> "" Then
						strcc = strcc & ";"
					End If
					strcc = strcc & rs3(0)
					rs3.movenext
				loop
				If strcc <> "" Then
					objMessage.cc = strcc
				End If
			Else
				if rs(2) <> "" then
					objMessage.cc = rs(2)			
				end If
			End If
			WScript.echo objMessage.from & ":" & objMessage.To &":" &objMessage.cc 
 			if (objMessage.cc <> "") then
				objMessage.cc = objMessage.cc 
			end if
			Wscript.Echo "SS : " & rs(14)
			if rs(7) <> "" then
				Set fso1 = CreateObject("Scripting.FileSystemObject")
				Set f1 = fso1.OpenTextFile(rs(7), 1)
				BodyText = f1.ReadAll()
				f1.Close
				Set f = Nothing
				Set fso = Nothing 
				if rs(14) = "N" then
					objMessage.TextBody = BodyText
				else
					BodyText.SetHtmlBody BodyText
				end if
				Wscript.Echo BodyText
			Else
						
				if Nomusaurio <> " " then
					if rs(14) = "N" then
					objMessage.TextBody = rs(4) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Nomusaurio & Chr(13) &  FonoUsuario & Chr(13) & "Centro de Servicios Financieros" & Chr(13) & ClienteUsuario & Chr(13) & formatnumber(nrut,0) & sdv
					else
						BodyText.SetHtmlBody rs(4) & "<br>" & "<br>" & "<br>" & "<br>" & Nomusaurio & "<br>" &  FonoUsuario & "<br>" & "Centro de Servicios Financieros" & "<br>" & ClienteUsuario & "<br>" & formatnumber(nrut,0) & sdv & "<br>" 
				end if
				else
					if rs(14) = "N" then
						objMessage.TextBody = rs(4) 
					else
						BodyText.SetHtmlBody rs(4)
					end if
				end if
			end if
			Set rs2 = cnnBase.Execute("select ltrim(rtrim(adjunto)), isnull(nkey_factura, -1)  from mail_adjuntos where nkey_mail = " & rs(8) )
			do while not rs2.eof
				if rs2(0) <> "" then
					objMessage.Addattachment rs2(0)
				end if
				upd = "update factura set factenviadamail = 'S' where nkey_factura = " & rs2(1)
				Set rsupd = cnnBase.Execute(upd)
				rs2.movenext
			loop

			If rs(6) <> "" Then
				objMessage.Addattachment DIR_OUT & rs(6)		
			End If
			
			rs2.close

			'==This section provides the configuration information for the remote SMTP server.
 
			objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
			'Name or IP of Remote SMTP Server
                        objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MAIL_SERVER

			
	
			'Type of authentication, NONE, Basic (Base64 encoded), NTLM
			objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic	
			'Your UserID on the SMTP server
			objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = rs1(0)

			'Your password on the SMTP server
			objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = rs1(1)



			'Server port (typically 25)
                        objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 1025 

			'Use SSL for the connection (False or True)
			objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

			'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
			objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

			objMessage.Configuration.Fields.Update
			objMessage.Send
Wscript.Echo  Err.Number 

 			'if Err.Number  <> 0 then
 		'		upd = "update mail set enviado = 'E', dfechaenvio = getdate() where nkey_mail = " & rs(8)
	'			Set rsupd = cnnBase.Execute(upd)
 	'		Else
			
			upd = "update mail set enviado = 'S', dfechaenvio = getdate() where nkey_mail = " & rs(8)
				Set rsupd = cnnBase.Execute(upd)
	'		End If


			rs1.close
		end if			
		rs.movenext
	loop
	rs.close
	exit sub

actualiza_mailerr:
	if rs(8) <> "" then
		upd = "update mail set enviado = 'B', dfechaenvio = getdate() where enviado in ('X','N') and nkey_mail = " & rs(8)
		Wscript.Echo upd
		Set rsupd = cnnBase.Execute(upd)
	End if
end sub

sub grabaTablaMail(de,para,referencia,texto,adjunto,iddestinos,cliente)
	dim rs 
	If de = ""  Then
		de = EMAIL_CALSER
	End If
	
	sql = "insert into mail (de,para,referencia,texto,adjunto,enviado,idorigen,tablaorigen,iddestinos,nkey_cliente) values ("
	sql = sql & "'" & de & "','" & para& "','" & referencia & "','" & texto & "','" & adjunto & "','N', 0, ' '," & iddestinos & "," & cliente & ")"
	Set rs = cnnBase.Execute(sql)

end sub


sub EnviaMailNewChk ()
	dim rs 
	dim rs1
	dim rs2
	dim rs3
	Dim fso1, f1
	dim bodytext
	dim Nomusaurio 
	dim fonousuario
	dim clienteusuario
	Dim strcc
	dim sql
	Dim nrut
	Dim sdv
	dim cantenviado
	dim nuevoformato
	Dim shell, pythonExe, scriptPath, registroId, comando
    Dim resultado
	Dim cid
	
    On Error resume next

	pythonExe = "python.exe" 
	scriptPath = "Z:\fas\pruebachil\pruebaimg.py" 'server BD

	cantenviado = 0
	sql = "select ltrim(rtrim(de)) ,para ,ltrim(rtrim(isnull(cc,''))),ltrim(rtrim(referencia)) , REPLACE(texto, '#48 ', CONVERT(VARCHAR(10), GETDATE(), 105)), idmail, isnull(adjunto,'') , archbody, nkey_mail, ltrim(rtrim(tablaorigen)), idorigen , isnull(iddestinos,0),nkey_cliente , enviado, isnull(html,'N'), isnull(texto2,''), isnull(masivo,'N') ,isnull(texto3,'') ,isnull(texto4,'') ,isnull(texto5,''), isnull(texto6,''), isnull(texto7,''),isnull(texto8,''), isnull(texto9,''), isnull(texto10,''), nkey_cliente, nkey_deudor , esnuevomail  from mail where enviado = 'X' and ((esnuevomail='S' and isnull(generoImagen,'N') ='S') or (esnuevomail = 'N')) " 
	'sql = sql & " union all"
	'sql = sql & " select ltrim(rtrim(de)) ,ltrim(rtrim(para)) ,ltrim(rtrim(isnull(cc,''))), ltrim(rtrim(referencia)),ltrim(rtrim(texto)), idmail, ltrim(rtrim(isnull(adjunto,''))), '', nkey_mail, ltrim(rtrim(tablaorigen)), idorigen , isnull(iddestinos,0),nkey_cliente,   enviado , isnull(html,'N'),  isnull(texto2,''), isnull(masivo,'N') ,isnull(texto3,'') ,isnull(texto4,'') ,isnull(texto5,''), isnull(texto6,''), isnull(texto7,''),isnull(texto8,''), isnull(texto9,''), isnull(texto10,''), nkey_cliente, nkey_deudor  from mail where enviado = 'N' order by isnull(masivo,'N') asc"
	Wscript.Echo sql
	Set rs = cnnBase.Execute(sql)
	strcc = ""
	do while not rs.eof
		Wscript.Echo "MAIL  " & rs(1)
		Nomusaurio =  " "
		strcc = ""
		fonousuario = " "
		clienteusuario = " "
		sdv = " "
		sdv = " "
		nrut = " "
		if rs(16) = "N" then
			if rs(9) = "GestionDeudor" then
				Set rs1 = cnnBase.Execute("select cliente.holding, usuarios.snombre, analista.anexo, cliente.nrutHolding, cliente.sdigitoHolding  from  cliente, gestiondeudor, usuarios, analista where analista.sloginanalista = usuarios.slogin and cliente.nkey_cliente = gestiondeudor.nkey_cliente and gestiondeudor.nkey_gestiondeudor = " & rs(10) & " and gestiondeudor.slogin  = usuarios.slogin and analista.activo = 'S'  and usuarios.activo = 'S'   ")
				if not rs1.eof then
					Nomusaurio =  rs1(1)
					fonousuario = "Fono (56-2) 2599-" & rs1(2)
					clienteusuario = rs1(0)
					nrut =  rs1(3)
					sdv =  " - " & rs1(4)
				end if
			end  if
		end if


		set mailman = CreateObject("Chilkat_9_5_0.MailMan")
		Wscript.Echo "MAIL  " 
	        success = mailman.UnlockComponent(CHILKAT_UNLOCK_CODE)
		if success  <> 0 then 
				Wscript.Echo "CHILKATT : " & success
		end if

		Wscript.Echo "MAIL  " & rs(0)
		set email = CreateObject("Chilkat_9_5_0.Email")
		Wscript.Echo rs(0)
		email.Subject = rs(3)
		maildesde = rs(0)
		strcc = ""
		
		
		if rs(0) <> "" then

			Set rs1 = cnnBase.Execute("select ltrim(rtrim(mailuserori)) as 'mailuserori', ltrim(rtrim(mailpassori )) as 'mailpassori' from cliente where mailorigen = '" & rs(0) & "'" )
			if rs1.eof then
				Set rs1 = cnnBase.Execute("select ltrim(rtrim(mailuserori)) as 'mailuserori', ltrim(rtrim(mailpassori )) as 'mailpassori' from cliente where mailorigen = 'servicioalcliente@debtcontrol.cl'" )
				maildesde = "servicioalcliente@debtcontrol.cl" 
			end if
			email.From = maildesde
			email.AddMultipleTo  rs(1)

			If rs(13) = "N" Then
				Set rs3 = cnnBase.Execute("select ltrim(rtrim(isnull(usuarios.smail,'" & EMAIL_ADM & "')))  from usuarios, usuario_mail where usuarios.slogin = usuario_mail.slogin and idproceso = " & rs(11) & " and usuario_mail.nkey_cliente in (0, " & rs(12) & ")")
				do while not rs3.eof
					If strcc <> "" Then
						strcc = strcc & ","
					End If
					strcc = strcc & rs3(0)
					rs3.movenext
				loop
				If strcc <> "" Then
					Wscript.Echo "RCCCC :" & strcc & ":"
					email.AddMultipleCC strcc
				End If
			Else
				if rs(2) <> "" then
					If strcc <> "" Then
						strcc = strcc & ","
					End If
					Wscript.Echo "CC :" & strcc & ":"
					strcc = strcc & rs(2)
				end If
				If strcc <> "" Then
					Wscript.Echo "RCCCC :" & strcc & ":"
					email.AddMultipleCC strcc
				End If

			End If
			
			registroId = rs(8)
			resultado = 1
			
			if rs(7) <> "" then
				Set fso1 = CreateObject("Scripting.FileSystemObject")
				Set f1 = fso1.OpenTextFile(rs(7), 1)
				BodyText = f1.ReadAll()
				f1.Close
				Set f = Nothing
				Set fso = Nothing 
					if rs(14) = "N" then
						email.Body = rs(4) & rs(15) & rs(17) & rs(18) & rs(19) & rs(20) & rs(21) & rs(22) & rs(23) & rs(24)
					else
						if rs(27) = "S" then
							resultado = fso.FileExists("Z:\Ejecutables\salida\temp\" & registroId & ".png")
							If resultado <> 0 Then
								cid = email.AddRelatedFile("Z:\Ejecutables\salida\temp\" & registroId & ".png")
								email.SetHtmlBody "<html><body>" &  "<img src=""cid:" & cid & """>" &  "</body></html>"
								email.SetHtmlBody "<html><body>" & _
											"<!--[if mso]>" & _
											"<table role=""presentation"" width=""1200"" cellpadding=""0"" cellspacing=""0"" border=""0"" align=""left""><tr><td>" & _
											"<![endif]-->" & _
											"<img src=""cid:" & cid & """ width=""1200"" style=""width:100%; max-width:1200px; height:auto; display:block; border:0;"">" & _
											"<!--[if mso]>" & _
											"</td></tr></table>" & _
											"<![endif]-->" & _
											"<div style=""font-size:9px; color:#b0b0b0; margin-top:6px;"">NKEY_MAIL" & rs(8) & "</div>" & _
											"</body></html>"
								email.SetHeaderField "Message-ID", "<nkeymail" & rs(8) & "." & Format(Now, "yyyymmddhhnnss") & "@dbtlatam.com>"

							end if
						else
							email.SetHtmlBody rs(4) & rs(15)  & rs(17) & rs(18) & rs(19)  & rs(20) & rs(21) & rs(22) & rs(23) & rs(24)
						end if
					end if
			Else
				if Nomusaurio <> " " then
					if rs(14) = "N" then
						email.Body = rs(4) & rs(15)  & rs(17) & rs(18) & rs(19)  & rs(20) & rs(21) & rs(22) & rs(23) & rs(24)
					else
						if rs(27) = "S" then
							resultado = fso.FileExists("Z:\Ejecutables\salida\temp\" & registroId & ".png")
							If resultado <> 0 Then
								cid = email.AddRelatedFile("Z:\Ejecutables\salida\temp\" & registroId & ".png")
								'email.SetHtmlBody "<html><body>" &  "<img src=""cid:" & cid & """>" &  "</body></html>"
								'email.SetHtmlBody "<html><body>" & "<table role=""presentation"" width=""700"" cellpadding=""0"" cellspacing=""0"" border=""0"">" &  "<tr><td><img src=""cid:" & cid & """ width=""100%"" style=""width:100%; max-width:700px; height:auto; display:block;""></td></tr>" & "</table>" & "</body></html>"								
								email.SetHtmlBody "<html><body>" & _
											"<!--[if mso]>" & _
											"<table role=""presentation"" width=""1200"" cellpadding=""0"" cellspacing=""0"" border=""0"" align=""left""><tr><td>" & _
											"<![endif]-->" & _
											"<img src=""cid:" & cid & """ width=""1200"" style=""width:100%; max-width:1200px; height:auto; display:block; border:0;"">" & _
											"<!--[if mso]>" & _
											"</td></tr></table>" & _
											"<![endif]-->" & _
											"<div style=""font-size:9px; color:#b0b0b0; margin-top:6px;"">NKEY_MAIL" & rs(8) & "</div>" & _
											"</body></html>"
								email.SetHeaderField "Message-ID", "<nkeymail" & rs(8) & "." & Format(Now, "yyyymmddhhnnss") & "@dbtlatam.com>"
							end if
						else
							email.SetHtmlBody rs(4) & rs(15)  & rs(17) & rs(18) & rs(19)  & rs(20) & rs(21) & rs(22) & rs(23) & rs(24)
						end if
					end if
				else
					if rs(14) = "N" then  
						email.Body = rs(4) & rs(15)  & rs(17) & rs(18) & rs(19)  & rs(20) & rs(21) & rs(22) & rs(23) & rs(24)
					else
						if rs(27) = "S" then
							resultado = fso.FileExists("Z:\Ejecutables\salida\temp\" & registroId & ".png")
							If resultado <> 0 Then
								cid = email.AddRelatedFile("Z:\Ejecutables\salida\temp\" & registroId & ".png")
								email.SetHtmlBody "<html><body>" & _
											"<!--[if mso]>" & _
											"<table role=""presentation"" width=""1200"" cellpadding=""0"" cellspacing=""0"" border=""0"" align=""left""><tr><td>" & _
											"<![endif]-->" & _
											"<img src=""cid:" & cid & """ width=""1200"" style=""width:100%; max-width:1200px; height:auto; display:block; border:0;"">" & _
											"<!--[if mso]>" & _
											"</td></tr></table>" & _
											"<![endif]-->" & _
											"<div style=""font-size:9px; color:#b0b0b0; margin-top:6px;"">NKEY_MAIL" & rs(8) & "</div>" & _
											"</body></html>"
								email.SetHeaderField "Message-ID", "<nkeymail" & rs(8) & "." & Format(Now, "yyyymmddhhnnss") & "@dbtlatam.com>"
							end if
						else
							email.SetHtmlBody rs(4) & rs(15)  & rs(17) & rs(18) & rs(19)  & rs(20) & rs(21) & rs(22) & rs(23) & rs(24)
						end if
					end if
				end if
			end if

			Set rs2 = cnnBase.Execute("select ltrim(rtrim(adjunto)), isnull(nkey_factura, -1)  from mail_adjuntos where nkey_mail = " & rs(8) )
			do while not rs2.eof
				if rs2(0) <> "" then
					email.AddFileAttachment  rs2(0)
				end if
				upd = "update factura set factenviadamail = 'S' where nkey_factura = " & rs2(1)
				Set rsupd = cnnBase.Execute(upd)
				rs2.movenext
			loop
			If rs(6) <> "" Then
				email.AddFileAttachment DIR_OUT & rs(6)		
			End If
			
			rs2.close

			if resultado <> 0 then
				mailman.SmtpHost = MAIL_SERVER
				
				mailman.SmtpUsername = rs1(0)
				mailman.SmtpPassword = rs1(1)
				WScript.echo rs1(0) & ";"  & rs1(1) & ":"
				
				mailman.mailport = 1110
				mailman.smtpport = 25

				mailman.SmtpSsl = 1
				mailman.StartTLS = 0
				mailman.SmtpAuthMethod = "NONE"
				rs1.close
				upd = "update mail set enviado = 'S', dfechaenvio = getdate() where nkey_mail = " & rs(8)
				Set rsupd = cnnBase.Execute(upd)
				Set rs1 = cnnBase.Execute("select nkey_mail from mail where enviado = 'S' and nkey_mail = " & rs(8) )

				if rs1.eof then
					upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor) values ('NoEnvMail'," & rs(8) & ",99,getdate()," & rs(25) & "," & success & ")"
					Set rsupd = cnnBase.Execute(upd)
					'upd = "update mail set enviado = 'X', dfechaenvio = NULL where nkey_mail = " & rs(8)
					'Set rsupd = cnnBase.Execute(upd)
				else
					success = mailman.SendEmail(email)
					If (success <> 1) Then
						WScript.echo mailman.LastErrorText
						upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor) values ('ErrEnvMail'," & rs(8) & ",99,getdate()," &rs(25)&"," & success & ")"
						Set rsupd = cnnBase.Execute(upd)
						
						'upd = "update mail set enviado = 'X', dfechaenvio = NULL where nkey_mail = " & rs(8)
						'Set rsupd = cnnBase.Execute(upd)
					else
						cantenviado = cantenviado + 1
						'upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor) values ('OKEnvMail'," & rs(8) & ",99,getdate(), 0 ," & success & ")"
						'Set rsupd = cnnBase.Execute(upd)
						'WScript.echo  upd 
						mailman.LastErrorText
					end if	
				end if
				if resultado = -1 then
					fso.DeleteFile "Z:\Ejecutables\salida\temp\" & registroId & ".png", True
				end if
				WScript.echo "BUENO " & cantenviado
			end if	
		end if			
		rs.movenext
	loop
	rs.close
	exit sub

actualiza_mailerr:
	if rs(8) <> "" then
		upd = "update mail set enviado = 'B', dfechaenvio = getdate() where enviado in ('X','N') and nkey_mail = " & rs(8)
		Wscript.Echo upd
		Set rsupd = cnnBase.Execute(upd)
	End if
end sub

sub grabaTablaMail(de,para,referencia,texto,adjunto,iddestinos,cliente)
	dim rs 
	If de = ""  Then
		de = EMAIL_CALSER
	End If
	
	sql = "insert into mail (de,para,referencia,texto,adjunto,enviado,idorigen,tablaorigen,iddestinos,nkey_cliente) values ("
	sql = sql & "'" & de & "','" & para& "','" & referencia & "','" & texto & "','" & adjunto & "','N', 0, ' '," & iddestinos & "," & cliente & ")"
	Set rs = cnnBase.Execute(sql)

end sub
