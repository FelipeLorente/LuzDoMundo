<!--#include file="include/conexao.asp"-->
<!--#include file="include/topo.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<style>
body {
  padding-top: 30px;
}
</style>

<%'PARAMETROS OBRIGATORIOS
wprograma  	  = "usuarios"
usu_cod	      = Request.QueryString("usu_cod")
altera	      = Request.QueryString("altera")
erro 		  = ""
	
'VERIFICA A SESSION
IF LEN(SESSION("usu_cod")) = 0 THEN
	REDIRECTPAGE(2)	
	response.End() 
END IF

'RECUPERA TIPO
IF LEN(TRIM(Request.Form("tipo")))>0 THEN
	tipo	  = Request.Form("tipo")   
ELSE
    tipo	  = Request.QueryString("tipo") 
END IF

'VERIFICA SE TEM TIPO
IF NOT ISEMPTY (tipo) THEN
 	
 	'COLOCA ERRO 
	erro = 2
	                      
	'''''''''''''''''''RECUPERA CAMPOS DO FORM
	usu_nome			=	Request.Form("usu_nome")
    usu_login			=	Request.Form("usu_login")
    usu_senha			=	Request.Form("usu_senha")
    usu_email			=	Request.Form("usu_email") 
	usu_status			=	Request.Form("usu_status")  
	
	'VERIFICA SE TEM STATUS
	IF LEN(TRIM(usu_status)) = 0 THEN
		usu_status = "A"
	END IF 
	
	'PADRONIZA LETRA MAIUSCULA
	usu_nome			= TRIM(UCASE(usu_nome))
	usu_login			= TRIM(UCASE(usu_login))
	usu_senha			= TRIM(UCASE(usu_senha))
	usu_email			= TRIM(UCASE(usu_email))	
	usu_status			= TRIM(UCASE(usu_status))	

	'CADASTRA NA TABELA AGENCIAS
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_ma_usuarios"
	cmd.CommandType = 4
	Set params = cmd.Parameters

	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)
	params.Append cmd.CreateParameter("@tipo", 3, 1, 0, numerosql(tipo))
	params.Append cmd.CreateParameter("@usu_cod", 3, 1, 0, numerosql(usu_cod))
	params.Append cmd.CreateParameter("@usu_nome", 200, 1, 70, usu_nome)
	params.Append cmd.CreateParameter("@usu_login", 200,1,50, usu_login)
	params.Append cmd.CreateParameter("@usu_senha",200,1,20, usu_senha)		
	params.Append cmd.CreateParameter("@usu_email",200,1,100, usu_email)
	params.Append cmd.CreateParameter("@usu_status",200,1,1, usu_status)	
		
	cmd.Execute	
	 
	'VERIFICA SE TEM RETORNO
	IF LEN(params("RETURN_VALUE")) > 0 AND params("RETURN_VALUE") > 0 then

		IF params("RETURN_VALUE") = 1 THEN
			'COLOCA ERRO 
			erro = 1
		END IF
	ELSE
		'VERIF SE ESTA ALTERANDO
		IF tipo = 2 THEN
			erro = 3
		ELSE
			'COLOCA ERRO 
			erro = 0
		END IF
    END IF
	
END IF

'VERIFICA SE EXISTE ALGUM DADO PARA SER REGATADO
IF LEN(usu_cod)>0 THEN
 	
	'CONSULTA IGREJA PELO COD
	SET QRY = conexao.execute("pr_le_usuarios 1, "&usu_cod&"")
	
	'VERIF SE É FINAL DE ARQ
	IF NOT QRY.EOF THEN
		
		'RECUPERA CAMPOS
		usu_cod				= QRY("usu_cod") 
		usu_nome			= QRY("usu_nome") 
		usu_login			= QRY("usu_login") 
		usu_senha			= QRY("usu_senha") 
		usu_email			= QRY("usu_email") 
		usu_status			= QRY("usu_status") 
		usu_dtCadastro		= QRY("usu_dtCadastro") 
		usu_dtAlteracao		= QRY("usu_dtAlteracao") 

		QRY.close			:	SET QRY 		= nothing
	END IF
	
END IF%>

<form action="<%=wprograma%>.asp?usu_cod=<%=usu_cod%>" name="forme" id="forme" method="post" onSubmit="return validaform(this)">
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>   
    <td>
        <label class="control-label">NOME <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="usu_nome" id="usu_nome" value="<%=usu_nome%>" maxlength="70" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite o nome!" style="width:320px;"> 
    </td>
	<td width="10"></td>
	<td>
        <label class="control-label">LOGIN <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="usu_login" id="usu_login" value="<%=usu_login%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite o login!" style="width:120px;"> 
    </td>
	<td width="10"></td>
	<td>
        <label class="control-label">SENHA <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="usu_senha" id="usu_senha" value="<%=usu_senha%>" maxlength="20" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite a senha!" style="width:120px;"> 
    </td>
	<td width="10"></td>
    <td>
        <label class="control-label">E-MAIL:</label><br>
        <input type="text" name="usu_email" id="usu_email" value="<%=usu_email%>" maxlength="100" class="form-control input-sm text-uppercase" parametro="email" requerido="no" msg="" style="width:400px;"> 
    </td>
    
    <%'VERIF SE É ALTERAÇAO
	IF altera = "sim" THEN%>
        <td width="5"></td>
        <td>
            <label class="control-label">STATUS:</label><br>                                                                 
            <select name="usu_status" id="usu_status" class="selectpicker" onChange="confirmaDesativar(this.value)" title="Selecione..." parametro="no" requerido="no" msg="">	                    			
			<option value="A" <%IF ucase(trim(usu_status)) = "A" THEN response.write "selected"%>>ATIVO </option>
            <option value="D" <%IF ucase(trim(usu_status)) = "D" THEN response.write "selected"%>>DESATIVADO </option>
            </select>
        </td>
    <%END IF%>
</tr>
</table>
<br>
<table cellpadding="2" cellspacing="0" border="0" class="formulario" align="center">
<tr>
    <td>
		<%IF LEN(usu_cod) =0 THEN%>
			<div class="col-xs-10 col-sm-4 col-md-4 col-lg-4" style="text-align:center">
				<input type="hidden" id="tipo" name="tipo" value="1">
				<input type="submit" name="salvar" id="salvar" value="SALVAR" class="btn btn-default">
			</div>
	   <%ELSE%>
			<div class="col-xs-10 col-sm-4 col-md-4 col-lg-4" style="text-align:center">
				<input type="hidden" id="tipo" name="tipo" value="2">
				<input type="submit" name="alterar" id="alterar" value="ALTERAR" class="btn btn-default">
			</div>
		<%END IF%>
    </td>
    <td width="35">
    	<div id="loading7" class="loader hidden" style="width:30px; height:30px"></div>
    </td>
</tr>
</table>		        
</form>
<br>
<table cellpadding="2" cellspacing="0" border="0" align="center" width="95%">
<tr>   
    <td align="center"><iframe id="listaUsuarios" name="listaArquivos" frameborder="0" width="100%" src="listaUsu.asp" style="height:320px"></iframe></td>
</tr>
</table>    
</body>

<script>

//pega o height do elemento e coloca na div
var height = screen.height;
document.getElementById('listaUsuarios').height = (height-550)

</script>

<script>
////////////////////////////////////// erro
<%IF erro = 0 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Usuário cadastrado com sucesso!')
	$('#avisoImagem').addClass('glyphicon-ok text-success');
	
	// foco no botao
	$('#aviso').ready(function(e) {
         $('#ok').focus();	
    });				
    
	//foca o campo quando apertar ok
	$("#ok").click(function(){
		location.href = '<%=wprograma%>.asp'													
	});		
	
<%END IF%>

<%IF erro = 1 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Usuário já cadastrado!')
	$('#avisoImagem').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	$('#aviso').ready(function(e) {
         $('#ok').focus();	
    });				
    
<%END IF%>
	
<%IF erro = 2 THEN%>

// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Erro ao cadastrar usuário, contate o suporte!')
	$('#avisoImagem').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	$('#aviso').ready(function(e) {
		 $('#ok').focus();	
	});				

<%END IF%>

<%IF erro = 3 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Usuário alterado com sucesso!')
	$('#avisoImagem').addClass('glyphicon-ok text-success');
	
	// foco no botao
	$('#aviso').ready(function(e) {
         $('#ok').focus();	
    });				
    
	//foca o campo quando apertar ok
	$("#ok").click(function(){
		location.href = '<%=wprograma%>.asp'													
	});		
	
<%END IF%>
</script>

<%'FECHA TODAS AS QRY'S E CONEXÕES COM O BANCO DE DADOS
conexao.close			:	SET conexao 		= nothing%>