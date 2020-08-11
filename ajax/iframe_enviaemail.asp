<!--#include file="../include/conexao.asp"-->

<%'RECUPERA PARAMETROS
tipo			= request.QueryString("tipo")
email			= request.QueryString("email")

'VERIFICA SE EXISTE LOGIN E SENHA
IF tipo = 1 THEN

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_le_esqueceuSenha"
	cmd.CommandType = 4
	Set params = cmd.Parameters
	
	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
	params.Append cmd.CreateParameter("@usu_email", 200, 1, 100, email)	

	SET QRY = cmd.EXECUTE
		
	IF NOT QRY.EOF THEN%>
    	<script>		
			// reseta mensagem
			parent.$('#avisoEnvioEmail').addClass('hidden');	
			
			//bloquea funcao
			parent.$('#emailEsqSenha').prop("disabled", true);	
			parent.$('#fecharEsqSenha').prop("disabled", true);		
		</script>
        
        <%'RECUPERA CAMPOS 		
		usu_nome				= Trim(QRY("usu_nome"))
		usu_login 				= Trim(QRY("usu_login"))
		usu_senha			 	= Trim(QRY("usu_senha"))
		usu_email	 			= Trim(QRY("usu_email"))		
			
		'VERIFICA SE EXISTE SENHA
		IF LEN(usu_login) > 0 AND LEN(usu_senha) > 0 THEN
		
			'ENVIA EMAIL			
			body = ""
			
			'corpo do email
			body = body&"<body>"

			body = body&"<table border='0' cellpadding='0' cellspacing='0' width='90%' style='font-size: 13px; font-family: Tahoma; color: #000000;'>"
			body = body&"<tr height='5'>"
			body = body&"	<td></td>"
			body = body&"</tr>"
			body = body&"<tr>"
			body = body&"	<td>Segue sua senha para acessar o sistema.</td>"
			body = body&"</tr>"
			body = body&"<tr height='5'>"   
			body = body&"	<td></td>"
			body = body&"</tr>"
			body = body&"<tr height='15'>"
			body = body&"	<td></td>"
			body = body&"</tr>"
			body = body&"<tr>"
			body = body&"	<td>Login de acesso: <b style='font-size: 14px;'> "&usu_login&" </b></td>"
			body = body&"</tr>"	
			dy = body&"<tr height='15'>"
			body = body&"	<td></td>"
			body = body&"</tr>"
			body = body&"<tr>"
			body = body&"	<td>Senha de acesso: <b style='font-size: 14px;'> "&usu_senha&" </b></td>"
			body = body&"</tr>"	
			body = body&"<tr height='15'>"
			body = body&"	<td colspan='5'></td>"
			body = body&"</tr>"
			body = body&"</table>"
			body = body&"<table border='0' cellpadding='0' cellspacing='0' width='90%' style='font-size: 13px; font-family: Tahoma; color: #000000;'>"
			body = body&"<tr>"
			body = body&"	<td>Atenciosamente, </td>"
			body = body&"</tr>"
			body = body&"<tr height='15'>"
			body = body&"	<td></td>"
			body = body&"</tr>"
			body = body&"<tr>"
			body = body&"	<td>Luz do Mundo</td>"
			body = body&"</tr>"
			body = body&"<tr height='15'>"
			body = body&"	<td></td>"
			body = body&"</tr>"
			body = body&"</table>"
			body = body&"</body>"

			'RETORNA 0 = ENVIADO | 1 = NAO ENVIADO 			
			strResultado = EnviaEmail("Luz do Mundo", "igrluzdomundo@hotmail.com",usu_nome,email,"","","HTML","SOLICITAÇÃO DE SENHA",body,"")	
							
			IF strResultado = 1 THEN%>
			
				<script>		
					// deixa botao padrao
					parent.$("#enviarEsqSenha").button('reset')	
					
					// mostra msg
					parent.$('#avisoEnvioEmail').removeClass('hidden');
					parent.$('#avisoEnvioEmail').text('Erro ao enviar e-mail, favor entrar em contato com o suporte!');
					
					
					//desbloquea funcao	
					parent.$('#emailEsqSenha').prop("disabled", false);	
					parent.$('#fecharEsqSenha').prop("disabled", false);	
					
				</script> 
				
			<%ELSE%>
			
				<script>						
					// deixa botao padrao
					parent.$("#enviarEsqSenha").button('reset')
								
					// fecha envia email					
					parent.$('#esqueceuSenhaEmail').modal('toggle')		
					
					//avisoImagem glyphicon-remove														
					
					// mostra aviso de erro	
					parent.$('#avisoImagem').removeClass('glyphicon glyphicon-remove')	
					parent.$('#avisoImagem').addClass('glyphicon glyphicon-ok')	
					parent.$('#avisoImagem').css('color' , 'green')	
					parent.$('#avisoClass').removeClass('modal-sm')
					parent.$('#avisoClass').addClass('modal-lg')	
								
					parent.$('#aviso').modal('show')						
					parent.$('#avisoTitulo').text('Aviso do Sistema')
					parent.$('#avisoDescricao').text('Sua senha foi enviada por e-mail!')							
					parent.$('#senha').focus()
					
					//desbloquea funcao
					parent.$('#emailEsqSenha').prop("disabled", false);	
					parent.$('#fecharEsqSenha').prop("disabled", false);
																				
				</script>
		  <%END IF%> 
  		<%ELSE%>
        	
            <script>		
				// deixa botao padrao
				parent.$("#enviarEsqSenha").button('reset')	
				
				// mostra msg
				parent.$('#avisoEnvioEmail').removeClass('hidden');
				parent.$('#avisoEnvioEmail').text('Login e/ou senha não registrada no sistema. Entre em contato com o suporte!');
				
				//desbloquea funcao	
				parent.$('#emailEsqSenha').prop("disabled", false);	
				parent.$('#fecharEsqSenha').prop("disabled", false);	
				
			</script>
            
    	<%END IF%>             
   
	<%	QRY.close			:	SET QRY 		= nothing
	ELSE%>
    
    	<script>		
			// deixa botao padrao
			parent.$("#enviarEsqSenha").button('reset')	
			
			// mostra msg
			parent.$('#avisoEnvioEmail').removeClass('hidden');
			parent.$('#avisoEnvioEmail').text('E-mail inválido. Entre em contato com o suporte!');
			
			//desbloquea funcao	
			parent.$('#emailEsqSenha').prop("disabled", false);	
			parent.$('#fecharEsqSenha').prop("disabled", false);	
			
		</script>    
	<%END IF
END IF%>

<!--#include file="../include/fechaConexao.asp"-->


