<!--#include file="include/conexao.asp"-->
<!--#include file="include/topo.asp"-->

<%'PEGA PARAMETRO DE ERRO
erro 	= request.QueryString("erro")
tipo	= request.Form("tipo")
login	= request.Form("login")
senha	= request.Form("senha")

'VERIFICA TIPO  
IF tipo = 1 THEN

	'VERIFICA SQL INJECTION
	login 	= sqlInjection(login)
	senha 	= sqlInjection(senha)				 	
	erro	= 0
		
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_le_login"
	cmd.CommandType = 4
	Set params = cmd.Parameters
	
	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)
	params.Append cmd.CreateParameter("@usu_login", 200, 1, 50, login)
	params.Append cmd.CreateParameter("@usu_senha", 200, 1, 50, senha)	
	
	SET QRY = cmd.EXECUTE				
	
	IF NOT QRY.EOF THEN	
		
		'CADASTRO FINALIZADO
		IF QRY("usu_status") <> "A" THEN
			erro = 3		
		END IF
		
		'VERIFICA ERRO
		IF erro = 0 THEN
			
			'CRIA SESSAO DO CODIGO / ETAPA
			SESSION("usu_cod") 				= QRY("usu_cod")
			SESSION("usu_nome") 			= QRY("usu_nome")
			
			'CADASTRO SALVO
			response.Redirect("inicial.asp")	
			response.End() 
			
		END IF
		
		QRY.close			:	SET QRY 		= nothing
	ELSE
		'ALIMENTA VAR ERRO
		erro = 1 
	END IF  
	
END IF%>
  
  <nav class="navbar navbar-inverse navbar-fixed-top" style="background:#003d7b">
    <div class="container">
      <div class="navbar-header">        
        <a class="navbar-brand" href="default.asp" style="color:#FFFFFF; font-weight:bold">SISTEMA - Luz do Mundo</a>        
      </div>    	
    </div>
  </nav>
  
    <div class="container">
	   <form name="forme" id="forme" action="default.asp" method="post" class="form-signin" onSubmit="return validaform(this)">
       
       <table cellpadding="0" cellspacing="0" border="0" align="center">
       <tr>
       		<td valign="top" align="center">
            	
               <table border="0" cellpadding="0" cellspacing="0">
               <tr>
                    <td> <img src="imagens/logoLuzDoMundo.jpg"></td>
               </tr>
               <tr height="30" align="center">
               		<img src="imagens/bullet.gif" border="0" width="1" height="1" >
               </tr>          
               <tr id="campologin">
                    <td>
                    	<label  class="control-label" for="login">Login</label><br>
                        <input type="text" name="login" id="login" value="<%=login%>" maxlength="50" onBlur="buscaLogin(this.value)" onKeyUp="buscaLogin(this.value)" onKeyDown="buscaLogin(this.value)" class="form-control" title="Digite o login!" parametro="no" requerido="yes" msg="Digite o login!">
                    </td>
               </tr>
               <tr height="5" align="center">
               		<img src="imagens/bullet.gif" border="0" width="1" height="1">
               </tr> 
               <tr id="camposenha" class="hidden">
                    <td>     
                    	<label  class="control-label" for="senha">Senha</label><br>              	                      
                        <input type="password" name="senha" id="senha" value="" maxlength="50" class="form-control" placeholder="Senha" parametro="no" requerido="yes" msg="Digite a senha!">
                    </td>
               </tr>  
               <tr height="10" align="center">
               		<img src="imagens/bullet.gif" border="0" width="1" height="1" >
               </tr>             
               <tr>
                    <td colspan="2" align="center" valign="middle">    
                    	<input type="submit" id="entrar" name="entrar" value="Entrar" style="background:#003d7b;font-weight:bold;font-size:17px" class="btn btn-md btn-primary btn-block">
                        <input type="hidden" id="tipo" name="tipo" value="1">   
                    </td>
               </tr>
               <!--<tr height="10" align="center">
               		<img src="imagens/bullet.gif" border="0" width="1" height="1" >
               </tr>           
               <tr>
                    <td align="center"> <a onClick="esqueceuSenha()" class="form-signin-heading" data-whatever="@mdo"><font style="color:#003d7b;font-weight:bold;font-size:13px;cursor:pointer">Esqueceu a Senha?</font></a></td>
               </tr> -->
               </table>
                             
     	</td>
     </tr>
     <tr height="10" align="center">
         <img src="imagens/bullet.gif" border="0" width="1" height="1" >
     </tr>  
     <tr>
     	<td>
            <div class="alert alert-danger hidden" role="alert" id="avisoForm">
                <strong></strong>
            </div>
        </td>
     </tr>
     </table> 
     </form>
	
    </div> 
  </body>
</html>

<!-- MODAL ENVIA EMAIL-->
<div class="modal fade form-horizontal"  id="esqueceuSenhaEmail" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog" style="width: 550px;" role="document">
        <div class="modal-content">
            <!-- Modal Header -->
            <div class="modal-header" style="background:#003d7b">                
                <h4 class="modal-title" id="myModalLabel" style="color:#FFFFFF;font-weight:bold;font-size:16px">
                    Esqueceu a Senha
                </h4>
            </div>
            
            <!-- Modal Body -->
            <div class="modal-body">
                
                <br>               
                    
                <div id="erroInputemailEsqSenha" class="form-group">
                    <label  class="col-sm-3 control-label" for="emailEsqSenha">E-mail</label>
                    <div class="col-sm-7">
                    	<input type="text" name="emailEsqSenha" id="emailEsqSenha" maxlength="150" class="form-control" placeholder="Email: exemplo@dominio.com.br" parametro="email" requerido="yes" msg="Digite o email!"/>
                    </div>             
                    
                </div>                 
              	
                <br>
                
            	<div class="alert alert-danger hidden " role="alert" id="avisoEnvioEmail">
                	<strong></strong>
                </div>                      
                                          
            </div>
            
            <!-- Modal Footer -->
            <div class="modal-footer"> 
                <input type="button" id="fecharEsqSenha" name="fecharEsqSenha" value="Fechar" class="btn btn-primary" data-dismiss="modal" aria-label="Close" style="background:#003d7b;"; onClick="$('#login').attr('onBlur', 'buscaLogin(this.value)');">
                <input type="button" id="enviarEsqSenha" name="enviarEsqSenha" value="Enviar" class="btn btn-primary" data-loading-text="Enviando Email..." style="background:#003d7b;" onClick="esqueceuSenhaEnviaEmail()">
            </div>    

        </div>
    </div>    
    
</div>

<iframe name="frameaux" id="frameaux" src="" width="1" height="1" frameborder="0"></iframe>

</body>
  
<script type="text/javascript">
////////////////////////////////////// erro
<%IF erro = 1 THEN%>

	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Login e/ou senha inv�lidos!')
	$('#avisoImagem').addClass('glyphicon-info-sign text-info');
		
	// foco no botao
	$('#aviso').on('shown.bs.modal', function(){
	  $('#ok').focus();			  
	});	
	
	//redireciona para prox etapa
	$("#ok").click(function(){		
		$('#senha').focus();		
	});	
		    
<%END IF%>

<%IF erro = 2 THEN%>

	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Sua sess�o expirou, por favor fa�a o acesso novamente!')
	$('#avisoImagem').addClass('glyphicon-info-sign text-info');
		
	// foco no botao
	$('#aviso').on('shown.bs.modal', function(){
	  $('#ok').focus();			  
	});		
		    
<%END IF%>

<%IF erro = 3 THEN%>

	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Usu�rio n�o ativo nesse sistema!')
	$('#avisoImagem').addClass('glyphicon-info-sign text-info');
		
	// foco no botao
	$('#aviso').on('shown.bs.modal', function(){
	  $('#ok').focus();			  
	});		
		    
<%END IF%>

////////////////////////////////////// form principal
/*focus*/
$("#login").focus();

function buscaLogin(login){	

	if(login.length > 0){
	
		enviaLogin(login)				
	}	
}

function enviaLogin(){

	if($('#login').val().replace(/ /, '').length > 0){

		frameaux.location = 'ajax/iframe_login.asp?tipo=1&login=' + $('#login').val().replace(/ /, '');
	}
}


// funcao que envia email do esqueceu senha
function esqueceuSenhaEnviaEmail(){
	
	// muda botao
	$("#enviarEsqSenha").button('loading')
	$('#erroInputemailEsqSenha').removeClass('has-error');
	
	//bloquea funcao do botao
	$('#enviarEsqSenha').prop("disabled", false);
	
	// aviso
	$('#avisoEnvioEmail').addClass('hidden');
	$('#avisoEnvioEmail').text('');
	
	// remove aviso do input
	$('#erroInputemailEsqSenha').removeClass('has-error');			
	
	if($('#emailEsqSenha').val().length == 0){
		
		// muda cor do input
		$('#erroInputemailEsqSenha').addClass('has-error');	
		 
		// mostra msg
		$('#avisoEnvioEmail').removeClass('hidden');
		$('#avisoEnvioEmail').text('Digite o E-mail!');
		$("#emailEsqSenha").focus();		
		
		// deixa botao padrao
		$("#enviarEsqSenha").button('reset')
		return false;	
	}
	
	//envia paramentros
	frameaux.location = 'ajax/iframe_enviaemail.asp?tipo=1&email=' + $("#emailEsqSenha").val();		
}

// esqueceu senha
function esqueceuSenha(){
    
	// tira onblur do campo login
	$('#login').attr("onBlur", "");
	
	// muda caracteristica do modal
	$('#esqueceuSenhaEmail').modal({backdrop: 'static', keyboard: false}) 
	 
	// abre modal
	$('#esqueceuSenhaEmail').modal('show')
	
	//bloquea funcao do botao
	$('#enviarEsqSenha').prop("disabled", false);
	
	// aviso
	$('#avisoEnvioEmail').addClass('hidden');
	$('#avisoEnvioEmail').text('');
	
	// remove aviso do input
	$('#erroInputemailEsqSenha').removeClass('has-error');	
	
	// limpa campos
	$("#emailEsqSenha").val('')
			
	// foca no campo email
	$("#emailEsqSenha").focus();		
	
}

</script>

<!--#include file="include/fechaConexao.asp"-->