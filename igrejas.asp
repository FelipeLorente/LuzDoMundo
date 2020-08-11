<!--#include file="include/conexao.asp"-->
<!--#include file="include/topo.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<style>
body {
  padding-top: 30px;
  overflow:hidden
}
div.dropdown-menu{
  max-height: 315px !important;
  overflow: hidden;
}
ul.dropdown-menu{
  max-height: 270px !important;
  overflow-y: auto;
}
</style>

<%'PARAMETROS OBRIGATORIOS
wprograma  	  = "igrejas"
igr_cod	      = Request.QueryString("igr_cod")
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
	igr_cnpj				=	Request.Form("igr_cnpj")
    igr_rSocial				=	Request.Form("igr_rSocial")
    igr_nFantasia			=	Request.Form("igr_nFantasia")
    igr_cep					=	Request.Form("igr_cep")
    igr_logradouro			=	Request.Form("igr_logradouro")
    igr_numero				=	Request.Form("igr_numero")
    igr_complemento			=	Request.Form("igr_complemento")
    igr_bairro				=	Request.Form("igr_bairro")
    igr_cidade				=	Request.Form("igr_cidade")
    igr_estado				=	Request.Form("igr_estado")
    igr_responsavel			=	Request.Form("igr_responsavel")
  	igr_telResponsavel		=	Request.Form("igr_telResponsavel")
	igr_internacional		=	Request.Form("igr_internacional")
	igr_status				=	Request.Form("igr_status")
	
	'VERIFICA SE TEM STATUS
	IF LEN(TRIM(igr_status)) = 0 THEN
		igr_status = "A"
	END IF
	
	'VERIFICA SE É INTERNACIONAL
	IF igr_internacional = "SIM" THEN
		igr_internacional = "S"
	ELSE
		igr_internacional = "N"
	END IF
	
	'PADRONIZA LETRA MAIUSCULA
	igr_rSocial				= TRIM(UCASE(igr_rSocial))
	igr_nFantasia			= TRIM(UCASE(igr_nFantasia))
	igr_logradouro			= TRIM(UCASE(igr_logradouro))
	igr_complemento			= TRIM(UCASE(igr_complemento))
	igr_bairro				= TRIM(UCASE(igr_bairro))
	igr_cidade				= TRIM(UCASE(igr_cidade))
	igr_estado				= TRIM(UCASE(igr_estado))
	igr_responsavel			= TRIM(UCASE(igr_responsavel))
	igr_status				= TRIM(UCASE(igr_status))
	
	'RETIRA CARACTERES
	igr_cnpj = retiraCaracteres(igr_cnpj)

	'CADASTRA NA TABELA AGENCIAS
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_ma_igrejas"
	cmd.CommandType = 4
	Set params = cmd.Parameters

	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)
	params.Append cmd.CreateParameter("@tipo", 3, 1, 0, numerosql(tipo))
	params.Append cmd.CreateParameter("@igr_cod", 3, 1, 0, numerosql(igr_cod))
	params.Append cmd.CreateParameter("@igr_cnpj", 200,1,14, stringsqlserver(igr_cnpj))
	params.Append cmd.CreateParameter("@igr_rSocial",200,1,100, stringsqlserver(igr_rSocial))		
	params.Append cmd.CreateParameter("@igr_nFantasia",200,1,60, stringsqlserver(igr_nFantasia))	
	params.Append cmd.CreateParameter("@igr_cep",200,1,20,stringsqlserver(igr_cep))
	params.Append cmd.CreateParameter("@igr_logradouro",200,1,100,stringsqlserver(igr_logradouro))
	params.Append cmd.CreateParameter("@igr_numero",3,1,0,numerosql(igr_numero))
	params.Append cmd.CreateParameter("@igr_complemento",200,1,20, stringsqlserver(igr_complemento))
	params.Append cmd.CreateParameter("@igr_bairro",200,1,50, stringsqlserver(igr_bairro))
	params.Append cmd.CreateParameter("@igr_cidade",200,1,50,stringsqlserver(igr_cidade))	
	params.Append cmd.CreateParameter("@igr_estado",200,1,50,stringsqlserver(igr_estado))	
	params.Append cmd.CreateParameter("@igr_status",200,1,1,stringsqlserver(igr_status))	 
	params.Append cmd.CreateParameter("@igr_responsavel",200,1,50,stringsqlserver(igr_responsavel))
	params.Append cmd.CreateParameter("@igr_telResponsavel",200,1,20,stringsqlserver(igr_telResponsavel))
	params.Append cmd.CreateParameter("@igr_internacional",200,1,1,stringsqlserver(igr_internacional))
		
	cmd.Execute	
	 
	'VERIFICA SE TEM RETORNO
	IF LEN(params("RETURN_VALUE")) > 0 AND params("RETURN_VALUE") > 0 then

		IF params("RETURN_VALUE") = 1 THEN
			'COLOCA ERRO 
			erro = 1
		END IF
	ELSE
		'VERIF SE ESTA ALTERANDO
		IF tipo = 1 THEN
			erro = 0
		END IF
		
		IF tipo = 2 THEN
			erro = 3
		END IF
		
		IF tipo = 3 THEN
			erro = 4
		END IF
		
    END IF
	
END IF

'VERIFICA SE EXISTE ALGUM DADO PARA SER REGATADO
IF LEN(igr_cod)>0 THEN
 	
	'CONSULTA IGREJA PELO COD
	SET QRY = conexao.execute("pr_le_igrejas 1, "&igr_cod&"")
	
	'VERIF SE É FINAL DE ARQ
	IF NOT QRY.EOF THEN
		
		'RECUPERA CAMPOS
		wstatus					= "SIM"
		igr_cod					= QRY("igr_cod") 
		igr_cnpj 				= QRY("igr_cnpj")
		igr_rSocial 			= QRY("igr_rSocial") 
		igr_nFantasia 			= QRY("igr_nFantasia") 
		igr_cep 				= QRY("igr_cep") 
		igr_logradouro 			= QRY("igr_logradouro") 
		igr_numero 				= QRY("igr_numero") 
		igr_complemento 		= QRY("igr_complemento") 
		igr_bairro 				= QRY("igr_bairro")
		igr_cidade 				= QRY("igr_cidade") 
		igr_estado 				= QRY("igr_estado") 
		igr_status 				= QRY("igr_status") 
		igr_responsavel 		= QRY("igr_responsavel") 
		igr_telResponsavel 		= QRY("igr_telResponsalvel")
		igr_internacional 		= QRY("igr_internacional") 		
		
		est_sigla = igr_estado
		
		QRY.close			:	SET QRY 		= nothing
	END IF
	
END IF

'FORMATA CNPJ
IF LEN(TRIM(igr_cnpj)) = 14 THEN
	igr_cnpj = FormataCNPJ(igr_cnpj)
ELSE
	igr_cnpj = ""
END IF%>

<iframe name="frameaux" id="frameaux" width="1" height="1" frameborder="0"></iframe>
	
<form action="<%=wprograma%>.asp?igr_cod=<%=igr_cod%>" name="forme" id="forme" method="post" onSubmit="return validaform(this)">
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>   
	<td>
        <label  class="control-label">CNPJ <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="igr_cnpj" id="igr_cnpj" readonly value="75.794.354/0001-54" maxlength="18" onBlur="mascara(this, cpf_mask); cnpjIncompleto(this.value)" onKeyUp="mascaraCnpj(this); mascara(this, cpf_mask)" onKeyDown="mascaraCnpj(this); mascara(this, cpf_mask)" onKeypress="return somenteNumeros(event);mascara(this, cpf_mask)" class="form-control" placeholder="CNPJ: 99.999.999/9999-99" title="Digite o cnpj!" parametro="no" requerido="yes" msg="Digite o Cnpj!" style="width:200px; height:30px">
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">RAZÃO SEDE<font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="igr_rSocial" id="igr_rSocial" value="<%=igr_rSocial%>" maxlength="100" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite o nome da igreja sede!" style="width:420px;"> 
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">IGREJA + NACIONALIDADE:<font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="igr_nFantasia" id="igr_nFantasia" value="<%=igr_nFantasia%>" maxlength="60" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite o nome da igreja!" style="width:350px;"> 
    </td>
    <td width="5"></td>	
    <td valign="bottom">
        <div class="checkbox" style="height:10px; min-height:10px; max-height:10px">
            <label>
                <input type="checkbox" id="igr_internacional" name="igr_internacional" onClick="verifCampo()" value="SIM" class="form-check-input position-static" <% IF TRIM(igr_internacional) = "S" THEN Response.Write "checked" END IF%> />
                <span class="cr" style="width:12px; height:12px"><i class="cr-icon glyphicon glyphicon-check" style="font-size:13px;"></i></span> <b>Igreja Internacional?</b>                                   
            </label>
        </div> 
    </td>
</tr>
</table>

<br>
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>     
	<td>
        <label class="control-label">CEP <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="igr_cep" id="igr_cep" value="<%=igr_cep%>" maxlength="9" class="form-control input-sm text-uppercase" onKeyUp="mascaraCep(this); buscaCep(this.value)" onKeyDown="mascaraCep(this);" onKeypress="return somenteNumeros(event); mascaraCep(this); buscaCep(this.value);" onBlur="buscaCep(this.value)" placeholder="99999-999" parametro="no" requerido="yes" msg="Digite o cep!" style="width:120px;"> 
    </td> 
    <td width="5"></td>         
    <td>
        <label class="control-label">LOGRADOURO <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="igr_logradouro" id="igr_logradouro" value="<%=igr_logradouro%>" maxlength="100" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite o endereço!" style="width:300px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">NÚMERO <font style="color:#FF0000;">*</font>:</label><br>                                                
        <input type="text" name="igr_numero" id="igr_numero" value="<%=igr_numero%>" maxlength="8" class="form-control input-sm" onKeypress="return somenteNumeros(event);" placeholder="999" parametro="no" requerido="yes" msg="Digite o numero!" style="width:80px;">
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">COMPLEMENTO:</label><br>                                                
        <input type="text" name="igr_complemento" id="igr_complemento" value="<%=igr_complemento%>" maxlength="20" class="form-control input-sm text-uppercase"  placeholder="CASA 1" parametro="no" requerido="no" msg="" style="width:110px;">
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">BAIRRO <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="igr_bairro" id="igr_bairro" value="<%=igr_bairro%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite o bairro!" style="width:150px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">CIDADE <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="igr_cidade" id="igr_cidade" value="<%=igr_cidade%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="yes" msg="Digite a cidade!" style="width:150px;">           
    </td>
    <td width="5"></td>
    
    <td>
       <label class="control-label">ESTADO <font style="color:#FF0000;">*</font>:</label><br>
       <div class="input-group dropdown">
            <input type="text" name="igr_estado" id="igr_estado" onBlur="validaCampoOri()" class="form-control countrycode dropdown-toggle" maxlength="50" value="" readonly="readonly" parametro="no" requerido="yes" msg="Selecione o Estado!" style="width:160px;text-transform:uppercase">
            <ul class="dropdown-menu" id="estado">
				<%'LISTA ESTADOS
				 SET QRY_EST = conexao.execute("pr_le_estados")
				 DO UNTIL QRY_EST.EOF%>
                 
                    <li><a href="#" data-value="<%=UCASE(QRY_EST("est_sigla") &" - "& QRY_EST("est_nome"))%>"><%=UCASE(QRY_EST("est_sigla") &" - "& QRY_EST("est_nome"))%></a></li>
                    
				<%QRY_EST.MOVENEXT
				LOOP
				QRY_EST.CLOSE			:	SET QRY_EST 		= nothing%>  
            </ul>
            <span role="button" class="input-group-addon dropdown-toggle" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false"><span class="caret"></span></span>
        </div>
         
	</td>
</tr>
</table>

<br>
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>
	<td colspan="10">
		<label class="control-label">NOME RESPONSAVEL PELA IGREJA:</label><br>
		<input type="text" name="igr_responsavel" id="igr_responsavel" value="<%=igr_responsavel%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:350px;"> 
	</td>
	<td width="5"></td>
	<td> 
		<label class="control-label">CELULAR:</label><br>           
		<input type="text" name="igr_telResponsavel" id="igr_telResponsavel" value="<%=igr_telResponsavel%>" placeholder="99 99999-9999" onKeyUp="mascaraCel2(this)" onKeyDown="mascaraCel2(this)" onKeyPress="return somenteNumeros(event)" onBlur="verifNum(this.value, this.id)" maxlength="13" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:126px;">
	</td>
    
    <%'VERIF SE É ALTERAÇAO
	IF altera = "sim" THEN%>
        <td width="5"></td>
        <td>
            <label class="control-label">STATUS:</label><br>                                                                 
            <select name="igr_status" id="igr_status" class="selectpicker" onChange="confirmaDesativar(this.value)" title="Selecione..." parametro="no" requerido="no" msg="">	                    			
			<option value="A" <%IF ucase(trim(igr_status)) = "A" THEN response.write "selected"%>>ATIVO </option>
            <option value="D" <%IF ucase(trim(igr_status)) = "D" THEN response.write "selected"%>>DESATIVADO </option>
            </select>
        </td>
    <%END IF%>
</tr>
</table>
<br>
<table cellpadding="2" cellspacing="0" border="0" class="formulario" align="center">
<tr>
    <td>
		<%IF LEN(igr_cod) =0 THEN%>
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

<table cellpadding="2" cellspacing="0" border="0" class="formulario" align="center" width="100%" height="280">
<tr>
	<td><iframe name="listaIgr" id="listaIgr" width="100%" height="270" src="listaIgr.asp" frameborder="0"></iframe></td>
</tr>
</table>
</body>


<script>

// faz função do select na lista nao ordenada estado
$(function() {
  $('#estado a').click(function() {
	console.log($(this).attr('data-value'));
	$(this).closest('.dropdown').find('input.countrycode')
	  .val($(this).attr('data-value'))	
  });
});

<%IF wstatus = "SIM" THEN
	'VERIFICA SE É INTERNACIONAL
	IF igr_internacional = "S" THEN%>
		// coloca valor no campo			
		$('#igr_estado').val('<%=igr_estado%>')		
		
		// chama funcao
		formatCampos(1)
	<%ELSE%>
	
		// coloca valor no campo			
		$('#igr_estado').val('<%=igr_estado%>')		
		
		// chama funcao
		formatCampos(2)
		
		//refresh no campo
		$(".selectpicker").selectpicker('refresh')
	<%END IF%>
<%END IF%>

//formata cep
function formatCampos(tipo){
	
	//1 = internacional / 2 = nacional
	if(tipo == 1){
		$("#igr_cep").attr('maxlength' , '15')
		$("#igr_cep").attr('onKeyUp' , '')
		$("#igr_cep").attr('onKeyDown' , '')
		$("#igr_cep").attr('onKeypress' , '')
		$("#igr_cep").attr('onBlur' , '')
		$("#igr_cep").attr('placeholder' , '###########')
		
		$('#igr_estado').prop('readonly', false);
		
		$("#igr_telResponsavel").attr('onKeyUp' , '')
		$("#igr_telResponsavel").attr('onKeyDown' , '')
		$("#igr_telResponsavel").attr('placeholder' , '###########')
		$("#igr_telResponsavel").attr('maxlength' , '20')
		$("#igr_telResponsavel").attr('onKeyPress' , '')
		$("#igr_telResponsavel").attr('onBlur' , '')
		
	}
	
	if(tipo == 2){
		
		$("#igr_cep").attr('maxlength' , '9')
		$("#igr_cep").attr('onKeyUp' , 'mascaraCep(this); buscaCep(this.value)')
		$("#igr_cep").attr('onKeyDown' , 'mascaraCep(this)')
		$("#igr_cep").attr('onKeypress' , 'return somenteNumeros(event); mascaraCep(this); buscaCep(this.value);')
		$("#igr_cep").attr('onBlur' , 'buscaCep(this.value)')
		$("#igr_cep").attr('placeholder' , '99999-999')
		
		$('#igr_estado').prop('readonly', true);
		
		$("#igr_telResponsavel").attr('onKeyUp' , 'mascaraCel2(this)')
		$("#igr_telResponsavel").attr('onKeyDown' , 'mascaraCel2(this)')
		$("#igr_telResponsavel").attr('placeholder' , '99 99999-9999')
		$("#igr_telResponsavel").attr('maxlength' , '13')
		$("#igr_telResponsavel").attr('onKeyPress' , 'return somenteNumeros(event)')
		$("#igr_telResponsavel").attr('onBlur' , 'verifNum(this.value, this.id)')

	}
	
}

// verifica se é internacional
function verifCampo(){
	
	if($('#igr_internacional').prop("checked") == true){
		$('#igr_internacional').val('SIM')
		$('#igr_estado').prop('readonly', false);
		
		// chama funcao
		formatCampos(1)
	}else{
		$('#igr_internacional').val('')
		$('#igr_estado').prop('readonly', true);
		
		// chama funcao
		formatCampos(2)
	}
	
	// limpa campos do endereco
	$('#igr_cep').val('')
	$('#igr_logradouro').val('')
	$('#igr_numero').val('')
	$('#igr_complemento').val('')
	$('#igr_bairro').val('')
	$('#igr_cidade').val('')
	$('#igr_estado').val('')
	
	$('#igr_cep').focus()
}

//verifica se o CNPJ esta completo
function cnpjIncompleto(valor){
	
	if(valor.length != 18){
		$('#igr_cnpj').val('')
		
		// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
		$('#aviso').modal({backdrop: 'static', keyboard: false}) 
		
		// mostra aviso de erro
		$('#aviso').modal('show')
		$('#avisoTitulo').text('Aviso do Sistema')
		$('#avisoDescricao').text('CNPJ incorreto!')
		$('#avisoImagem').addClass('glyphicon-remove text-danger');
		
		// foco no botao
		$('#aviso').ready(function(e) {
			 $('#ok').focus();	
		});
	}
}

// confirma se quer desativar igreja 
function confirmaDesativar(valor){

	if(valor == "D"){
		
		// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
		$('#aviso').modal({backdrop: 'static', keyboard: false}) 
		
		// mostra aviso de erro
		$('#aviso').modal('show')
		$('#avisoTitulo').text('Aviso do Sistema')
		$('#avisoDescricao').text('Ao desativar esta igreja TODOS os seus membros serão desativados automáticamente!')
		$('#avisoImagem').addClass('glyphicon-remove text-danger');
		
		// foco no botao
		$('#aviso').ready(function(e) {
			 $('#ok').focus();	
		});												
	}
}

//ajax para buscar cep nacional
function buscaCep(cep){
	
	if(cep.length == 9){	 
		frameaux.location = 'ajax/iframe_geral.asp?tipo=1&cep=' + cep
	}	 
}

// monta mascara do cnpj quando sai do campo form principal
function mascara(o,f){

	v_obj=o
	v_fun=f
	setTimeout("execmascara()",1)
}

function execmascara(){
	v_obj.value=v_fun(v_obj.value)
	
	// verifica se tem 18 caracteres para liberar campos senha
	if(v_obj.value.length == 18){
		enviaCnpj()				
	}	
}

/////////////////////////////// funcao generica
// mascara cnpj
function cpf_mask(v){

	v=v.replace(/\D/g,"")                 //Remove tudo o que não é dígito
	v=v.replace(/(\d{2})(\d)/,"$1.$2")    //Coloca ponto entre o terceiro e o quarto dígitos
	v=v.replace(/(\d{3})(\d)/,"$1.$2")    //Coloca ponto entre o setimo e o oitava dígitos
	v=v.replace(/(\d{3})(\d)/,"$1/$2")   //Coloca ponto entre o decimoprimeiro e o decimosegundo dígitos
	v=v.replace(/(\d{4})(\d)/,"$1-$2")   //Coloca ponto entre o decimoprimeiro e o decimosegundo dígitos
	
	return v
}

function enviaCnpj(){

	if($('#igr_cnpj').val().length == 18){
					
		frameaux.location = 'ajax/iframe_cnpj.asp?cnpj=' + $("#igr_cnpj").val();
	}
}

<%IF trim(altera) = "sim" THEN%>
	//bloqueia campos	
	$('#igr_cnpj').attr('disabled','disabled');
	
<%END IF%>
</script>

<script>
////////////////////////////////////// erro
<%IF erro = 0 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Igreja cadastrada com sucesso!')
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
	$('#avisoDescricao').text('Igreja já cadastrada!')
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
	$('#avisoDescricao').text('Erro ao cadastrar igreja, contate o suporte!')
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
	$('#avisoDescricao').text('Igreja alterada com sucesso!')
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

<%IF erro = 4 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Igreja e Irmãos excluídos com sucesso!')
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