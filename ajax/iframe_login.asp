<!--#include file="../include/conexao.asp"-->

<%'RECUPERA PARAMETROS
tipo			= request.QueryString("tipo")
login			= request.QueryString("login")
senha			= request.QueryString("senha")%>

<script>
// controle especifico do campo cnpj
// deixa controle padrao
parent.$('#avisoForm').addClass('hidden');
parent.$('#avisoForm').text('');
parent.$('#campologin').removeClass('has-error');

</script>

<%'VERIFICA LOGIN
IF tipo = 1 THEN
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_le_login"
	cmd.CommandType = 4
	Set params = cmd.Parameters
	
	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
	params.Append cmd.CreateParameter("@usu_login", 200, 1, 50, login)
	params.Append cmd.CreateParameter("@usu_senha", 200, 1, 50, "")	
	
	set QRY = cmd.Execute
	
	IF NOT QRY.EOF THEN%>
    	<script>		
			//libera campo senha				
			parent.$('#camposenha').removeClass('hidden')	
			parent.$('#senha').attr('requerido', 'yes')		
			parent.$('#senha').focus()
		</script>
        			
	<%	QRY.close			:	SET QRY 		= nothing
	ELSE%>
    	<script>		
			//libera campo senha
			parent.$('#login').focus()		
			parent.$('#senha').attr('requerido', 'no')
			parent.$('#camposenha').addClass('hidden')			
		</script>    
	<%END IF
END IF%>

<!--#include file="../include/fechaConexao.asp"-->
