set fso = CreateObject("Scripting.FileSystemObject")
Set cnnBase = CreateObject("ADODB.Connection")
Cnnbase.ConnectionTimeout = 0
Cnnbase.CommandTimeout = 0
ScriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
FileName = fso.BuildPath(ScriptDir,"globaldebtnet.vbs")
WScript.Echo Filename
set f = fso.OpenTextFile(FileName,1)
s = f.ReadAll()
ExecuteGlobal s
f.Close

' Config local con credenciales (mailsecrets.vbs no se sube a git)
SecretsFile = fso.BuildPath(ScriptDir,"mailsecrets.vbs")
set fSecrets = fso.OpenTextFile(SecretsFile,1)
ExecuteGlobal fSecrets.ReadAll()
fSecrets.Close

dim Tarde

cnnBase.ConnectionString = DSN_DBT
cnnBase.Open
call LeeMailCliente()


Sub LeeMailCliente()
	dim rs 
	dim rsBus
	dim rsBus2
	dim rsIns
	dim rsbusnkey
	Dim rsBusCon
	Dim rsBusAna
	dim rsupd
	dim x 
	dim loginanalista
	dim posx
	dim posy
	dim posz
	dim nomarch
	Dim forwardEmail
	Dim i
	dim numAttach
	dim ind
	dim nkey_mail
	dim posNkeyMail
                                        
    On Error Resume Next
 
	sql = "select distinct 0 as nkey_cliente, ltrim(rtrim(mailorigen)),ltrim(rtrim(mailuserori)),ltrim(rtrim(mailpassori)) from cliente where mailorigen is not null and bcuentacorriente <> 0 "
	sql =  sql &  " union  "
	sql =  sql & "select distinct -99 as nkey_cliente, ltrim(rtrim('" & REICH_MAIL_USER & "')),ltrim(rtrim('" & REICH_POP_USER & "')),ltrim(rtrim('" & REICH_POP_PASS & "')) from cliente where nkey_cliente = 527  "
	sql =  sql &  " union  "
	sql =  sql &   "select distinct -98 as nkey_cliente, ltrim(rtrim('" & TRANSVE_MAIL_USER & "')),ltrim(rtrim('" & TRANSVE_POP_USER & "')),ltrim(rtrim('" & TRANSVE_POP_PASS & "')) from cliente where nkey_cliente = 2939  "
	sql =  sql &  " union  "
	sql =  sql &   "select distinct -97 as nkey_cliente, ltrim(rtrim('" & CANON_MAIL_USER & "')),ltrim(rtrim('" & CANON_POP_USER & "')),ltrim(rtrim('" & CANON_POP_PASS & "')) from cliente where nkey_cliente = 240 "

	' WScript.echo sql
	Set rs = cnnBase.Execute(sql)

	set mailman = CreateObject("Chilkat_9_5_0.MailMan")

   	success = mailman.UnlockComponent(CHILKAT_UNLOCK_CODE)
	
	If (success <> 1) Then
	    WScript.echo mailman.LastErrorText
    	    WScript.Quit
	End If



	do while not rs.eof
'		WScript.echo rs(2) & ":" & rs(3)  & ":"

		mailman.MailHost = "127.0.0.1"
		mailman.mailport = 1110

		mailman.StartTLS  = 0
		WScript.echo "version" & mailman.TlsVersion 
		mailman.PopUsername = rs(2)
		mailman.PopPassword = rs(3)	
		mailman.ImmediateDelete = 1
		mailman.MaxCount = 30
' 		'mailman.pop3ssl = true
		WScript.echo rs(2) & ":" & rs(3)  & ":"
		
		WScript.echo  mailman.getmailboxcount() 

		Set bundle = mailman.GetAllHeaders(10)


		If (bundle Is Nothing ) Then
			WScript.echo  " Error Mail entrantes " 
			WScript.echo mailman.LastErrorText
		else
			WScript.echo  "Mail entrantes " & bundle.MessageCount
		end if
		total = bundle.MessageCount - 1
		if total > 10 then
			total = 10
		else
			total = bundle.MessageCount - 1
		end if

		textoBuscar = "NKEY_MAIL"
		
		For i = 0 To total 
		    con_error  = 0
		    ' email is a Chilkat.Email2
		    WScript.echo  "Voy en " & i & " de un total de " & total
		    Set email_new  = bundle.GetEmail(i)
		    Set email  = mailman.GetfullEmail(email_new)

			' --- DIAGNOSTICO TEMPORAL ---
			WScript.Echo "----- DIAG mensaje " & i & " -----"
			WScript.Echo "Asunto: " & email.Subject
			WScript.Echo "HasHtmlBody(): " & email.HasHtmlBody()
			WScript.Echo "Len(GetHtmlBody()): " & Len(email.GetHtmlBody())
			WScript.Echo "Len(GetPlainTextBody()): " & Len(email.GetPlainTextBody())
			WScript.Echo "NumAttachments: " & email.NumAttachments
			WScript.Echo "In-Reply-To: [" & email.GetHeaderField("In-Reply-To") & "]"
			WScript.Echo "References:  [" & email.GetHeaderField("References") & "]"
			WScript.Echo "ContentType (Mime): " & Left(email.GetMime(), 400)
			WScript.Echo "----- FIN DIAG -----"
			' --- FIN DIAGNOSTICO TEMPORAL ---

			If email.HasHtmlBody() Then
				' Extraemos el HTML puro con todas sus etiquetas intactas
				textmail = Trim(email.GetHtmlBody())
			Else
				' Si el correo fue enviado puramente como texto, usamos el plan B
				textmail = Trim(email.GetPlainTextBody())
			End If
		    'textmail = trim(email.GetMime())
			
			
			WScript.echo  len(textmail)
			
		    posx =  InStr(1,textmail,"KC-")
		    posy =  InStr(1,textmail,"-KD-")
		    posz =  InStr(1,textmail,"-KT-")
			posh =  InStr(1,textmail,"-KH-")
		    WScript.echo  "KTKC en " & i &  posx  & "  " & posy & "  " & posz & " " & posh
			varnkey_cli = 0
			varnkey_deu = 0
		    if (posx > 0 and posy > posx and posz > posy ) then
				varnkey_cli = mid(textmail, posx+3,posy-posx-3)
				varnkey_deu = mid(textmail, posy+4,posz-posy-4)
				if posh > 0 then
					varnkey_ges = mid(textmail, posz+4,posh-posz-4)
				end if
				WScript.echo varnkey_cli & " " & varnkey_deu & " " & varnkey_ges
			end  if
			' --- Metodo 1 (intento rapido): header In-Reply-To/References ---
			' Nota: en este ambiente Exchange reescribe el Message-ID en envios
			' anonimos, asi que normalmente no va a matchear; se deja como intento
			' barato por si en algun momento deja de reescribirse.
			nkey_mail = 0
			posNkeyMail = 0
			inReplyTo = Trim(email.GetHeaderField("In-Reply-To"))
			if inReplyTo = "" then
				inReplyTo = Trim(email.GetHeaderField("References"))
			end if
			posHeaderKey = InStr(1, inReplyTo, "nkeymail", vbTextCompare)
			if posHeaderKey > 0 then
				nkey_mail = ExtraeDigitosTras(inReplyTo, "nkeymail")
				if nkey_mail > 0 then
					posNkeyMail = posHeaderKey
					WScript.Echo "Llave recuperada desde In-Reply-To: " & nkey_mail
				end if
			end if

			' --- Metodo 2 (principal): marcador NKEY_MAIL visible en el body ---
			' Se extraen los digitos que siguen al marcador sin importar que venga
			' despues (funciona tanto con el comentario viejo <!--NKEY_MAIL123-->
			' como con el <div> chico y visible del template nuevo).
			if nkey_mail = 0 then
				posNkeyMail =  InStr(1,textmail,textoBuscar,vbTextCompare)
				if posNkeyMail > 0 then
					nkey_mail = ExtraeDigitosTras(textmail, textoBuscar)
					WScript.Echo "Llave recuperada desde el body: " & posNkeyMail
					WScript.Echo "Llave recuperada desde el body: " & nkey_mail
				end if
			end if
			
			if posNkeyMail > 0 then
				sql = "select  top 1 cliente.nkey_cliente, ltrim(rtrim(cliente.pathadjuntosvcor)), mail.nkey_deudor from cliente, mail  where mail.nkey_mail = " & nkey_mail & " and mail.nkey_cliente = cliente.nkey_cliente and cliente.bcuentacorriente <> 0"
			else
				if varnkey_ges > 0 then
					sql = "select  distinct cliente.nkey_cliente, ltrim(rtrim(cliente.pathadjuntosvcor)), gestiondeudor.nkey_deudor from cliente, gestiondeudor  where gestiondeudor.ngestionagrupa = " & varnkey_ges & " and gestiondeudor.nkey_cliente = cliente.nkey_cliente and cliente.bcuentacorriente <> 0"
				else 
					if varnkey_cli = 0 then
						sql = "select top 1 nkey_cliente, ltrim(rtrim(pathadjuntosvcor)), cliente.nkey_cliente from cliente where ltrim(rtrim(mailorigen)) = '"& rs(1) & "' and bcuentacorriente <> 0"
					else
						sql = "select nkey_cliente, ltrim(rtrim(pathadjuntosvcor)), cliente.nkey_cliente from cliente where nkey_cliente = "	 & varnkey_cli
					end if
				end if
			end if
		    
			set rsbusnkey = cnnBase.Execute(sql)
			do while not rsbusnkey.eof 
				WScript.echo  rsbusnkey(0)
				varnkey_cli = rsbusnkey(0)
				if posNkeyMail > 0 then
					varnkey_deu = rsbusnkey(2)
				end if
				textmail = trim(email.GetPlainTextBody())
				textmail = replace(textmail,"'","") 
				if (len(textmail) > 950) then
					textmail = mid(textmail,1,950)
				end if
				if (varnkey_deu = 0 ) then
					sql = "select contactosdeudor.nkey_deudor, contactosdeudor.nKey_ContactoDeudor from contactosdeudor join cliente on (cliente.nkey_cliente = " &  varnkey_cli & " )  join codigodeudor  on "
					sql = sql & " (cliente.nkey_cliente = codigodeudor.nkey_cliente and codigodeudor.nkey_deudor = contactosdeudor.nkey_deudor) where cliente.nkey_cliente = contactosdeudor.nkey_cliente  and " 
					sql = sql & " contactosdeudor.sEmail = '" & email.Fromaddress & "' and isnull(contactosdeudor.activo,'S') = 'S' "
				else
					sql = "select codigodeudor.nkey_deudor, isnull(contactosdeudor.nKey_ContactoDeudor,1) from codigodeudor  join cliente on (cliente.nkey_cliente = " &  varnkey_cli & " ) left join contactosdeudor  on (contactosdeudor.nkey_cliente = codigodeudor.nkey_cliente and codigodeudor.nkey_deudor =  contactosdeudor.nkey_deudor) where cliente.nkey_cliente = codigodeudor.nkey_cliente  and contactosdeudor.sEmail  = '" & email.Fromaddress & "' and isnull(contactosdeudor.activo,'S') = 'S' and codigodeudor.nkey_deudor = " &  varnkey_deu 				
				end if

				Set rsBusCon = cnnBase.Execute(sql)

				if rsBusCon.eof then  'No existe el mail registrado al deudor cliente

					estado = "R"

					Call grabaTablaMailRecep(email.Fromaddress,rs(1),email.Subject,textmail," ",0,varnkey_cli,email.uidl,estado,0," ",0)
					n = email.NumAttachments
					if n > 0 then
						success = email.SaveAllAttachments(rsbusnkey(1) & "adj-entrada\" )  '"u:\Ejecutables\salida\adj-entrada\"
					end if
					for j = 0 to n -1


						sql  = "select nkey_mail  from mail where nkey_cliente = " & varnkey_cli & " and enviado = '" & estado & "' and archbody = '" & email.uidl & "'"

						Set rsbus2 = cnnBase.Execute(sql)

						if not rsbus2.eof then
							sql = "insert into mail_adjuntos (nkey_mail, adjunto) values (" & rsbus2(0) & " , '" & rsbusnkey(1) & "adj-entrada\"  & email.getattachmentfilename(j) & "')"
							Set rsIns = cnnBase.Execute(sql)
						end if
					Next 
				Else
					do while not rsBusCon.eof 		    	
						if email.Fromaddress <> rs(1) then
							sql = "select analista.sloginanalista from servicio join analista on (analista.nkey_analista = servicio.nkey_analista) where nkey_cliente = " & varnkey_cli & " and nkey_deudor = " & rsBusCon(0)

							Set rsBusAna = cnnBase.Execute(sql)

							if rsBusAna.eof then
								loginanalista = "mail"
							else
								loginanalista = rsBusAna(0)
							end if
							estado = "T"
							sql = "select top 1 * from gestiondeudor where nkey_cliente = " & varnkey_cli & " and nkey_deudor = " & rsBusCon(0) & " and sobjetivo = '" & email.uidl & "' order by nkey_gestiondeudor desc "
							Set rsBus2 = cnnBase.Execute(sql)
							sql = "insert into gestiondeudor (nKey_Cliente,nKey_Deudor,nKey_Contacto,sTipoGestion,sDeudor,sanalista,sObservacion,dFechaGestion,dHoraGestion,nKey_Factura, sLogin,sobjetivo )"
							sql = sql & " values ( "							
							sql = sql & varnkey_cli & ", " & rsBusCon(0)  & ", " & rsBusCon(1)  & ",'Escrita Entrante', "
							if rsBus2.eof then
								sql = sql & " 'Respuesta Mail'"
							else
								sql = sql & " 'Mail Entrante'"
							end if
							sql = sql & " , '" & email.Subject & "','" & textmail & "',  getdate(), getdate(),0,'" & loginanalista  & "','" & email.uidl & "')"
							Set rsBus = cnnBase.Execute(sql)
							sql = "select top 1 nkey_gestiondeudor from gestiondeudor where nkey_cliente = " & varnkey_cli & " and nkey_deudor = " & rsBusCon(0) & " and nkey_contacto = " & rsBusCon(1) & " and sobjetivo = '" & email.uidl & "' order by nkey_gestiondeudor desc "
							Set rsBus = cnnBase.Execute(sql)
							nkey_gestion = rsBus(0)
							
							Call grabaTablaMailRecep(email.Fromaddress,rs(1),email.Subject,textmail," ",0,varnkey_cli,email.uidl,estado,nkey_gestion,"GestionDeudor",rsBusCon(0))
							n = email.NumAttachments
							if n > 0 then
								success = email.SaveAllAttachments(rsbusnkey(1) & "adj-entrada\" )  
							end if
							for j = 0 to n -1
								sql  = "select nkey_mail  from mail where nkey_cliente = " & varnkey_cli & " and enviado = '" & estado & "' and archbody = '" & email.uidl & "'"
								Set rsbus2 = cnnBase.Execute(sql)
								if not rsbus2.eof then
									sql = "insert into mail_adjuntos (nkey_mail, adjunto) values (" & rsbus2(0) & " , '" & rsbusnkey(1) & "adj-entrada\"  & email.getattachmentfilename(j) & "')"
									Set rsIns = cnnBase.Execute(sql)
								end if
							Next 
						end if	
						rsBusCon.movenext 
					loop
				End If
				rsbusnkey.movenext 
			loop

			posx =  InStr(1,ucase(email.Fromaddress),ucase("@paperless.pe"))

			if posx > 0 then
				upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor) values ('MailPaper',99,99,getdate()," & varnkey_cli & ", 0)"
				Set rsupd = cnnBase.Execute(upd)
				WScript.echo  upd
				'Es Facturador electronico Peru
				sql = "select  direcXMLFP, trim(pathsdocscaneados)+trim(direcarchivos)+'\' from cliente where ncod = 4096"
				Set rsbus2 = cnnBase.Execute(sql)
				if not rsbus2.eof then
					
					for n = 0 to email.NumAttachments - 1
						if  InStr(1,Ucase(email.GetAttachmentFilename(n)),ucase(".xml"))  > 0 then
							success = email.SaveAttachedFile(n, rsbus2(0))
						end if
						if  InStr(1,Ucase(email.GetAttachmentFilename(n)),ucase(".pdf"))  > 0 then
							xini = InStr(1,email.GetAttachmentFilename(n),"-")
							if xini > 0  then
								xfin = InStr(xini,email.GetAttachmentFilename(n),"_")
							end if
							if xfin > (xini+1) then
								success = email.SaveAttachedFile(n, rsbus2(1))
								Set objFSO = CreateObject("Scripting.FileSystemObject")
								objFSO.MoveFile rsbus2(1) & email.GetAttachmentFilename(n) , rsbus2(1) & "P" & mid(email.GetAttachmentFilename(n),xini+1,xfin-xini-1) & ".pdf"
							end if
						end if
					next 
				end if
			end if

			posx = 0 
			posx =  InStr(1,ucase(email.Fromaddress),ucase("facturacion@reich.cl"))



			if rs(0) = -99 and posx > 0  then
				varnkey_cli = 527
				'Es Facturador electronico reich
				sql = "select  direcXMLFP, trim(pathsdocscaneados)+trim(direcarchivos)+'\' from cliente where nkey_cliente = " & varnkey_cli
				Set rsbus2 = cnnBase.Execute(sql)
				if not rsbus2.eof then
					for n = 0 to email.NumAttachments - 1
upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor, observaciones) values ('MailReich',99,99,getdate()," &  varnkey_cli & ", 0,'" &  Ucase(email.GetAttachmentFilename(n)) & "' )"
						
						Set rsupd = cnnBase.Execute(upd)
						
						posx =  InStr(1,ucase(email.Subject),ucase(" T")) 
						tipodoc = "00"
						numdoc="0"
						WScript.Echo posx
						if posx > 0 then
							tipodoc = mid(ucase(email.Subject),posx+2,2)
						end if

						posx =  InStr(1,ucase(email.Subject),ucase(" F")) 
						WScript.Echo posx
						if posx > 0 then
							numdoc = mid(ucase(email.Subject),posx+2)
						end if
						If tipodoc= "33" Then
							tipodoc = "FE"
						end if

						If tipodoc= "61" Then
						        tipodoc = "NC"
						end if
						If tipodoc= "56" Then
						        tipodoc = "ND"
    						end if
						If tipodoc = "34" Then
        						tipodoc = "FE"
    						end if
						If tipodoc = "01" Then
						        tipodoc = "FP"
						end if
						WScript.Echo rsbus2(1) & tipodoc  &  numdoc & ".pdf"

						if tipodoc <> "00" then
							if  InStr(1,Ucase(email.GetAttachmentFilename(n)),ucase(".xml"))  > 0 then
								success = email.SaveAttachedFile(n, rsbus2(0))
							end if
							if  InStr(1,Ucase(email.GetAttachmentFilename(n)),ucase(".pdf"))  > 0 then
								success = email.SaveAttachedFile(n, rsbus2(1))									   	
								if len(email.GetAttachmentFilename(n)) > 11  then
									success = email.SaveAttachedFile(n, rsbus2(1))
									Set objFSO = CreateObject("Scripting.FileSystemObject")
									upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor, observaciones) values ('MailReich2',99,99,getdate()," &  varnkey_cli & ", 0,'" &  tipodoc  &  numdoc & ".pdf" & "' )"
									objFSO.MoveFile rsbus2(1) & email.GetAttachmentFilename(n) , rsbus2(1) & tipodoc  &  numdoc & ".pdf"
									WScript.Echo "MUEVE" & "  " & rsbus2(1) & email.GetAttachmentFilename(n) & " ---- " & rsbus2(1) & tipodoc  &  numdoc & ".pdf"
								end if
							end if
						end if
					next 
				end if
			end if


			posx = 0 
			posx =  InStr(1,trim(ucase(email.Fromaddress)),ucase("odoo@transve.cl"))
			if posx <= 0 then
				posx =  InStr(1,trim(ucase(email.Fromaddress)),ucase("odoo2@transve.cl"))
			end if
			if rs(0) = -98 and posx > 0  then

						
				Set rsupd = cnnBase.Execute(upd)	
				varnkey_cli = 2939
upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor, observaciones) values ('PreMailTransve',99,99,getdate()," &  varnkey_cli & ", 0,'" &  Ucase(ucase(email.Subject)) & "' )"
						Set rsupd = cnnBase.Execute(upd)
				'Es Facturador electronico transve
				sql = "select  direcXMLFP, trim(pathsdocscaneados)+trim(direcarchivos)+'\' from cliente where nkey_cliente = " & varnkey_cli
				Set rsbus2 = cnnBase.Execute(sql)
				WScript.Echo "CANTI ADJ " & email.GetAttachmentFilename(n) & email.NumAttachments


				if not rsbus2.eof then
					for n = 0 to email.NumAttachments - 1
						WScript.Echo  email.GetAttachmentFilename(n)
upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor, observaciones) values ('MailTransve',99,99,getdate()," &  varnkey_cli & ", 0,'" &  Ucase(email.GetAttachmentFilename(n)) & "' )"
						
						Set rsupd = cnnBase.Execute(upd)
						if n = 0 Then 
							upd = "insert into mail (de, para, referencia, texto, enviado, tablaorigen, dfechaenvio, nkey_cliente, masivo, esnuevomail) values ('facturas.transve@dbtlatam.com', 'quironixapp@gmail.com', '" & ucase(email.Subject) & "', 'Reenvio Factura', 'X', 'Nada', getdate(), '2939','N','N') "   
							Set rsupd = cnnBase.Execute(upd)
							WScript.Echo upd
							upd = "select isnull(max(nkey_mail),0) from mail where nkey_cliente = 2939 and tablaorigen = 'Nada'"
							Set rsBusnkey = cnnBase.Execute(upd)
							nkey_mail = rsBusnkey(0)
WScript.Echo nkey_mail

						end if


						
						
						referencia  =  ucase(email.Subject)

						posx =  InStr(1,ucase(referencia),ucase(" - ")) 
						if posx > 0 then
							codigodeudor = mid(referencia,posx+3)
						end if

						tipodoc = "00"
						posx =  InStr(1,ucase(referencia),ucase(" (Ref ")) 
						if posx > 0 then
							tipodoc = ucase(mid(referencia,posx+6,3))
						end if
						
						posy =  InStr(posx,ucase(referencia),ucase(")")) 

						if posy >0 then
							numdoc = mid(referencia,posx+10,posy - posx -10 )
						end if

						If tipodoc= "FAC" Then
							tipodoc = "FE"
						end if

						If tipodoc= "N/C" Then
						        tipodoc = "NC"
						end if
						If tipodoc= "N/D" Then
						        tipodoc = "ND"
    						end if

						WScript.Echo  referencia
						WScript.Echo rsbus2(1) & tipodoc  &  numdoc & ".pdf" & " Codeu " & codigodeudor 

						if tipodoc <> "00" then
							if  InStr(1,Ucase(email.GetAttachmentFilename(n)),ucase(".xml"))  > 0 then
								WScript.Echo "ENTRE XML " 
								success = email.SetAttachmentFilename(n, codigodeudor & "_" & tipodoc  &  numdoc & ".xml")
								if success <> 1 then
									WScript.Echo "ENTRE XML 1 " 
									con_error  = 1
								end if

								success = email.SaveAttachedFile(n, rsbus2(0))
								if success <> 1 then 
									WScript.Echo "ENTRE XML 2 " 
									con_error  = 1
								end if

								Set fsoexi = CreateObject("Scripting.FileSystemObject")
								filenameexi = rsbus2(0) & "\" &  codigodeudor & "_" & tipodoc  &  numdoc & ".xml"
								
								fileExistsexi = fso.FileExists(filenameexi)

								If fileExistsexi Then
								  'WScript.Echo "File exists!"
									upd = "insert into mail_adjuntos  (nkey_mail, adjunto) values (" & nkey_mail & " , '" &  filenameexi & "')"
WScript.Echo upd
									Set rsupd = cnnBase.Execute(upd)
								Else
								  WScript.Echo "ENTRE XML 3 " 
								  WScript.Echo "NO EXISTE " & filenameexi
									con_error  = 1
								End If


							end if
							if  InStr(1,Ucase(email.GetAttachmentFilename(n)),ucase(".pdf"))  > 0 then
								success = email.SetAttachmentFilename(n, tipodoc  &  numdoc & ".pdf")
								if success <> 1 then
									WScript.Echo "ENTRE PDF 1 " 
									con_error  = 1
								end if

								success = email.SaveAttachedFile(n, rsbus2(1))	
								if success <> 1 then
									WScript.Echo "ENTRE PDF 2 " 
									con_error  = 1
								end if
								   	
								if len(email.GetAttachmentFilename(n)) > 11  then
									success = email.SaveAttachedFile(n, rsbus2(1))
								if success <> 1 then 
									WScript.Echo "ENTRE PDF 3 " 
									con_error  = 1
								end if

								Set fsoexi = CreateObject("Scripting.FileSystemObject")
								filenameexi = rsbus2(1) & tipodoc  &  numdoc & ".pdf"
								
								fileExistsexi = fso.FileExists(filenameexi)

								If fileExistsexi Then
								  'WScript.Echo "File exists!"
									upd = "insert into mail_adjuntos  (nkey_mail, adjunto) values (" & nkey_mail & " , '" &  filenameexi & "')"
WScript.Echo upd
									Set rsupd = cnnBase.Execute(upd)

								Else
									WScript.Echo "NO EXISTE " & filenameexi
									WScript.Echo "ENTRE PDF 4 " 
									con_error  = 1
								End If


								Set objFSO = CreateObject("Scripting.FileSystemObject")
									upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor, observaciones) values ('MailTransve2',99,99,getdate()," &  varnkey_cli & 	", 0,'" &  tipodoc  &  numdoc & ".pdf" & "' )"
' objFSO.MoveFile rsbus2(1) & email.GetAttachmentFilename(n) , rsbus2(1) & tipodoc  &  numdoc & ".pdf"
 WScript.Echo "MUEVE" & "  " & rsbus2(1) & "XXXXXX" & email.GetAttachmentFilename(n) & " ---- " & rsbus2(1) & tipodoc  &  numdoc & ".pdf"													   	

								end if
							end if
						end if
					next 
				end if
			end if



			posx = 0 
			posx =  InStr(1,trim(ucase(email.Fromaddress)),ucase("@canon.cl"))


			if rs(0) = -97 and posx > 0  then
						
				Set rsupd = cnnBase.Execute(upd)	
				varnkey_cli = 240
upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor, observaciones) values ('ConciliaMail',99,99,getdate()," &  varnkey_cli & ", 0,'" &  Ucase(ucase(email.Subject)) & "' )"
				Set rsupd = cnnBase.Execute(upd)
				'Es Facturador electronico transve
				
				sql = "select  trim(pathsdocscaneados)+trim(nombrecorto)+'\'+'ArchivosConcilia\',trim(direcarchivos)+'\'+'ArchivosConcilia\' from cliente where nkey_cliente = " & varnkey_cli

				Set rsbus2 = cnnBase.Execute(sql)
				WScript.Echo "CANTI ADJ " & email.GetAttachmentFilename(n) & email.NumAttachments
				if not rsbus2.eof then
					for n = 0 to email.NumAttachments - 1
						WScript.Echo  email.GetAttachmentFilename(n)
upd = "insert into auditoria  (tabla, pantalla, usuario, fecha, nkey_cliente, nkey_deudor, observaciones) values ('ConciliaMail',99,99,getdate()," &  varnkey_cli & ", 0,'" &  Ucase(email.GetAttachmentFilename(n)) & "' )"
						
						Set rsupd = cnnBase.Execute(upd)
						
						
						referencia  =  ucase(email.Subject)

						if  InStr(1,Ucase(email.GetAttachmentFilename(n)),ucase(".xl"))  > 0 then
							WScript.Echo "ENTRE XLS " 
							
							success = email.SetAttachmentFilename(n, Replace(email.GetAttachmentFilename(n), " ", ""))
							if success <> 1 then
								WScript.Echo "ENTRE XLS 1 " 
								con_error  = 1
							end if

							success = email.SaveAttachedFile(n, rsbus2(0))
							if success <> 1 then 
								WScript.Echo "ENTRE XLS 2 " 
								con_error  = 1
							end if

							Set fsoexi = CreateObject("Scripting.FileSystemObject")
							filenameexi = rsbus2(0) & "\" &  Replace(email.GetAttachmentFilename(n), " ", "")
							
							fileExistsexi = fso.FileExists(filenameexi)

							If fileExistsexi Then
							  'WScript.Echo "File exists!"
							Else
								  WScript.Echo "ENTRE XLS 3 " 
								  WScript.Echo "NO GRABE ADJUNTO  " & filenameexi
								  con_error  = 1
							End If
						end if
					next 
				end if
			end if

			if con_error  = 0 then
				success = mailman.DeleteEmail(email_new)
				WScript.Echo "BORRE " 
			end if


		Next
		mailman.Pop3EndSession 
		rs.movenext
	loop
	rs.close
	DisplayErrorInfo
end sub


sub grabaTablaMailRecep(de,para,referencia,texto,adjunto,iddestinos,cliente, uid, estado,idorigen, tablaorigen,key_deudor)
	dim rstt 
	
	On Error Resume Next

	sql  = "select *  from mail where nkey_cliente = " & Cliente & " and enviado = '" & estado & "' and archbody = '" & uid & "'"
'			WScript.echo sql
	Set rstt = cnnBase.Execute(sql)
	if rstt.eof then
		sql = "insert into mail (de,para,referencia,texto,adjunto,enviado,idorigen,tablaorigen,iddestinos,nkey_cliente, archbody,nkey_deudor) values ("
		sql = sql & "'" & de & "','" & para& "','" & referencia & "','" & texto & "','" & adjunto & "','" & estado & "', " & idorigen & ", '" & tablaorigen & "'," & iddestinos & "," & cliente & ",'" & uid & "'," & key_deudor & ")"
		Set rstt = cnnBase.Execute(sql)
	end if
	DisplayErrorInfo
end sub

sub DisplayErrorInfo()
  For Each errorObject In rs.ActiveConnection.Errors
    WScript.script "Description: " & errorObject.Description & Chr(10) & Chr(13) & _
           "Number:      " & Hex(errorObject.Number)
  Next
end sub

' Busca "marcador" dentro de "texto" y devuelve los digitos que vienen
' inmediatamente despues, sin importar que caracter los corte (espacio,
' "<", "-->", etc). Devuelve 0 si no encuentra nada.
Function ExtraeDigitosTras(texto, marcador)
	Dim pos, ini, valor, ch
	ExtraeDigitosTras = 0
	pos = InStr(1, texto, marcador, vbTextCompare)
	if pos = 0 then Exit Function
	ini = pos + Len(marcador)
	valor = ""
	Do While ini <= Len(texto)
		ch = Mid(texto, ini, 1)
		if ch >= "0" and ch <= "9" then
			valor = valor & ch
			ini = ini + 1
		else
			Exit Do
		end if
	Loop
	if valor <> "" then ExtraeDigitosTras = CLng(valor)
End Function
