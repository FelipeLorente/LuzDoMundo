<!--#include file="include/conexao.asp"-->
<!--#include file="include/topo.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<style>
body {
  padding-top: 30px;
}
div.dropdown-menu{
  max-height: 315px !important;
  overflow: hidden;
}
ul.dropdown-menu{
  max-height: 270px !important;
  overflow-y: auto;
}

.blue-textarea textarea.md-textarea:focus:not([readonly]) {
  border-bottom: 1px solid #3B5998;
  box-shadow: 0 1px 0 0 #3B5998;
}

.active-blue-textarea.md-form textarea.md-textarea:focus:not([readonly])+label {
  color: #3B5998;
}
</style>

<%'PARAMETROS OBRIGATORIOS
wprograma  	  = "irmaos"
irm_cod	      = Request.QueryString("irm_cod")
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
	igr_cod					= Request.Form("igr_cod")
	irm_nome				= Request.Form("irm_nome")
	irm_nPai				= Request.Form("irm_nPai")
	irm_nMae				= Request.Form("irm_nMae")
	irm_cep					= Request.Form("irm_cep")
	irm_logradouro			= Request.Form("irm_logradouro")
	irm_numero				= Request.Form("irm_numero")
	irm_complemento			= Request.Form("irm_complemento")
	irm_bairro				= Request.Form("irm_bairro")
	irm_cidade				= Request.Form("irm_cidade")
	irm_estado				= Request.Form("irm_estado")
	irm_tel1				= Request.Form("irm_tel1")
	irm_tel2				= Request.Form("irm_tel2")
	irm_tpCel1				= Request.Form("irm_tpCel1")
	irm_cel1				= Request.Form("irm_cel1")
	irm_tpCel2				= Request.Form("irm_tpCel2")
	irm_cel2				= Request.Form("irm_cel2")
	irm_nacionalidade		= Request.Form("irm_nacionalidade")
	irm_estCivil			= Request.Form("irm_estCivil")
	irm_nConjegue			= Request.Form("irm_nConjegue")
	irm_qtdFilhos			= Request.Form("irm_qtdFilhos")
	irm_rg					= Request.Form("irm_rg")
	irm_cpf					= Request.Form("irm_cpf")
	irm_crente				= Request.Form("irm_crente")
	irm_btEspiritos			= Request.Form("irm_btEspiritos")
	irm_dtBtEspiritos		= Request.Form("irm_dtBtEspiritos")
	irm_btAguas				= Request.Form("irm_btAguas")
	irm_dtBtAguas			= Request.Form("irm_dtBtAguas")
	irm_denominacaoBatismo	= Request.Form("irm_denominacaoBatismo")
	irm_CidEstPais			= Request.Form("irm_CidEstPais")
	irm_dtIngresso			= Request.Form("irm_dtIngresso")
	irm_funcIgreja			= Request.Form("irm_funcIgreja")
	irm_dtNascimento		= Request.Form("irm_dtNascimento")
	irm_obs					= Request.Form("irm_obs")		
	
	'PADRONIZA LETRA MAIUSCULA
	irm_nome				= TRIM(UCASE(irm_nome))
	irm_nPai				= TRIM(UCASE(irm_nPai))
	irm_nMae				= TRIM(UCASE(irm_nMae))
	irm_cep					= TRIM(UCASE(irm_cep))
	irm_logradouro			= TRIM(UCASE(irm_logradouro))			
	irm_numero				= TRIM(UCASE(irm_numero))
	irm_complemento			= TRIM(UCASE(irm_complemento))
	irm_bairro				= TRIM(UCASE(irm_bairro))
	irm_cidade				= TRIM(UCASE(irm_cidade))
	irm_estado				= TRIM(UCASE(irm_estado))
	irm_tpCel1				= TRIM(UCASE(irm_tpCel1))
	irm_tpCel2				= TRIM(UCASE(irm_tpCel2))
	irm_tel1				= TRIM(UCASE(irm_tel1))
	irm_tel2				= TRIM(UCASE(irm_tel2))
	irm_cel1				= TRIM(UCASE(irm_cel1))
	irm_cel2				= TRIM(UCASE(irm_cel2))
	irm_nacionalidade		= TRIM(UCASE(irm_nacionalidade))
	irm_estCivil			= TRIM(UCASE(irm_estCivil))
	irm_nConjegue			= TRIM(UCASE(irm_nConjegue))
	irm_qtdFilhos			= TRIM(UCASE(irm_qtdFilhos))
	irm_rg					= TRIM(UCASE(irm_rg))
	irm_cpf					= TRIM(UCASE(irm_cpf))
	irm_crente				= TRIM(UCASE(irm_crente))
	irm_btEspiritos			= TRIM(UCASE(irm_btEspiritos))
	irm_btAguas				= TRIM(UCASE(irm_btAguas))
	irm_denominacaoBatismo	= TRIM(UCASE(irm_denominacaoBatismo))
	irm_CidEstPais			= TRIM(UCASE(irm_CidEstPais))
	irm_funcIgreja			= TRIM(UCASE(irm_funcIgreja))
	irm_dtNascimento		= TRIM(UCASE(irm_dtNascimento))
	irm_obs					= TRIM(UCASE(irm_obs))	

	'CADASTRA NA TABELA AGENCIAS
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_ma_irmaos"
	cmd.CommandType = 4
	Set params = cmd.Parameters

	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)
	params.Append cmd.CreateParameter("@tipo", 3, 1, 0, numerosql(tipo))
	params.Append cmd.CreateParameter("@igr_cod", 3, 1, 0, numerosql(igr_cod))
	params.Append cmd.CreateParameter("@irm_cod", 3, 1, 0, numerosql(irm_cod))
	params.Append cmd.CreateParameter("@irm_nome", 200,1,70, stringsqlserver(irm_nome))	
	params.Append cmd.CreateParameter("@irm_nPai", 200,1,70, stringsqlserver(irm_nPai))
	params.Append cmd.CreateParameter("@irm_nMae", 200,1,70, stringsqlserver(irm_nMae))
	params.Append cmd.CreateParameter("@irm_cep", 200,1,20, stringsqlserver(irm_cep))
	params.Append cmd.CreateParameter("@irm_logradouro", 200,1,100, stringsqlserver(irm_logradouro))
	params.Append cmd.CreateParameter("@irm_numero", 200,1,20, stringsqlserver(irm_numero))
	params.Append cmd.CreateParameter("@irm_complemento", 200,1,20, stringsqlserver(irm_complemento))
	params.Append cmd.CreateParameter("@irm_bairro", 200,1,50, stringsqlserver(irm_bairro))
	params.Append cmd.CreateParameter("@irm_cidade", 200,1,50, stringsqlserver(irm_cidade))
	params.Append cmd.CreateParameter("@irm_estado", 200,1,50, stringsqlserver(irm_estado))
	params.Append cmd.CreateParameter("@irm_tel1", 200,1,20, stringsqlserver(irm_tel1))
	params.Append cmd.CreateParameter("@irm_tel2", 200,1,20, stringsqlserver(irm_tel2))
	params.Append cmd.CreateParameter("@irm_tpCel1", 200,1,20, stringsqlserver(irm_tpCel1))
	params.Append cmd.CreateParameter("@irm_cel1", 200,1,20, stringsqlserver(irm_cel1))
	params.Append cmd.CreateParameter("@irm_tpCel2", 200,1,20, stringsqlserver(irm_tpCel2))
	params.Append cmd.CreateParameter("@irm_cel2", 200,1,20, stringsqlserver(irm_cel2))	
	params.Append cmd.CreateParameter("@irm_nacionalidade", 200,1,50, stringsqlserver(irm_nacionalidade))
	params.Append cmd.CreateParameter("@irm_estCivil", 200,1,50, stringsqlserver(irm_estCivil))
	params.Append cmd.CreateParameter("@irm_nConjegue", 200,1,70, stringsqlserver(irm_nConjegue))
	params.Append cmd.CreateParameter("@irm_qtdFilhos", 3, 1, 0, numerosql(irm_qtdFilhos))
	params.Append cmd.CreateParameter("@irm_rg", 200,1,20, stringsqlserver(irm_rg))
	params.Append cmd.CreateParameter("@irm_cpf", 200,1,20, stringsqlserver(irm_cpf))
	params.Append cmd.CreateParameter("@irm_crente", 200,1,1, stringsqlserver(irm_crente))
	params.Append cmd.CreateParameter("@irm_btEspiritos", 200,1,1, stringsqlserver(irm_btEspiritos))
	params.Append cmd.CreateParameter("@irm_dtBtEspiritos", 200,1,30, stringsqlserver(irm_dtBtEspiritos))	
	params.Append cmd.CreateParameter("@irm_btAguas", 200,1,1, stringsqlserver(irm_btAguas))
	params.Append cmd.CreateParameter("@irm_dtBtAguas", 200,1,30, stringsqlserver(irm_dtBtAguas))
	params.Append cmd.CreateParameter("@irm_denominacaoBatismo", 200,1,70, stringsqlserver(irm_denominacaoBatismo))
	params.Append cmd.CreateParameter("@irm_CidEstPais", 200,1,50, stringsqlserver(irm_CidEstPais))
	params.Append cmd.CreateParameter("@irm_dtIngresso", 200,1,30, stringsqlserver(irm_dtIngresso))	
	params.Append cmd.CreateParameter("@irm_funcIgreja", 200,1,50, stringsqlserver(irm_funcIgreja))
	params.Append cmd.CreateParameter("@irm_dtNascimento", 200,1,50, stringsqlserver(irm_dtNascimento))	
	params.Append cmd.CreateParameter("@irm_obs", 201,1,-1, stringsqlserver(irm_obs))					
		
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
IF LEN(irm_cod)>0 THEN
 	
	'CONSULTA IGREJA PELO COD
	SET QRY = conexao.execute("pr_le_irmaos 1, "&irm_cod&"")
	
	'VERIF SE É FINAL DE ARQ
	IF NOT QRY.EOF THEN
		
		'RECUPERA CAMPOS
		wstatus					= "SIM"
		igr_cod 				= QRY("igr_cod")
		igr_internacional		= QRY("igr_internacional")
		irm_nome 				= QRY("irm_nome")
		irm_nPai 				= QRY("irm_nPai")
		irm_nMae 				= QRY("irm_nMae")
		irm_cep 				= QRY("irm_cep")
		irm_logradouro 			= QRY("irm_logradouro")
		irm_numero 				= QRY("irm_numero")
		irm_complemento 		= QRY("irm_complemento")
		irm_bairro				= QRY("irm_bairro")
		irm_cidade 				= QRY("irm_cidade")
		irm_estado 				= QRY("irm_estado")
		irm_tel1 				= QRY("irm_tel1")
		irm_tel2 				= QRY("irm_tel2")
		irm_tpCel1				= QRY("irm_tpCel1")
		irm_cel1 				= QRY("irm_cel1")
		irm_tpCel2 				= QRY("irm_tpCel2")
		irm_cel2 				= QRY("irm_cel2")
		irm_nacionalidade 		= QRY("irm_nacionalidade")
		irm_estCivil 			= QRY("irm_estCivil")
		irm_nConjegue 			= QRY("irm_nConjegue")
		irm_qtdFilhos 			= QRY("irm_qtdFilhos")
		irm_rg 					= QRY("irm_rg")
		irm_cpf 				= QRY("irm_cpf")
		irm_crente 				= QRY("irm_crente")
		irm_btEspiritos 		= QRY("irm_btEspiritos")
		irm_dtBtEspiritos 		= QRY("irm_dtBtEspiritos")
		irm_btAguas 			= QRY("irm_btAguas")
		irm_dtBtAguas 			= QRY("irm_dtBtAguas")
		irm_denominacaoBatismo 	= QRY("irm_denominacaoBatismo")
		irm_CidEstPais 			= QRY("irm_CidEstPais")
		irm_dtIngresso 			= QRY("irm_dtIngresso")
		irm_funcIgreja			= QRY("irm_funcIgreja")
		irm_dtNascimento 		= QRY("irm_dtNascimento")
		irm_obs 				= QRY("irm_obs")	
		
		' FORMATA DATA
		IF LEN(TRIM(irm_dtNascimento))>0 THEN
			irm_dtNascimento = formataData(irm_dtNascimento)
		END IF
				
		QRY.close			:	SET QRY 		= nothing
	END IF
	
END IF

'VERIF SE TEM IRMAOS CADASTRADOS
SET QRY_LIS = conexao.execute("pr_le_irmaos 2, NULL")

IF NOT QRY_LIS.EOF THEN
	wstatusLista = "yes"
END IF

'FORMATA CNPJ
IF LEN(TRIM(igr_cnpj)) = 14 THEN
	igr_cnpj = FormataCNPJ(igr_cnpj)
ELSE
	igr_cnpj = ""
END IF%>

<iframe name="frameaux" id="frameaux" width="1" height="1" frameborder="0"></iframe>

<%IF wstatusLista = "yes" THEN%>
    <!--BOTAO PARA ABRIR A LISTA-->
    <div class="" style="position:fixed; top:10px; right:0px">
        <button type="button" class="btn btn-default btn-sm" onClick="abreModal()">
          <span class="glyphicon glyphicon-th-list" aria-hidden="true"></span> Lista
        </button>
    </div>
<%END IF%>
	
<form action="<%=wprograma%>.asp?irm_cod=<%=irm_cod%>" name="forme" id="forme" method="post" onSubmit="return validaform(this)">
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>   
	<td>
		<label class="control-label">IGREJA <font style="color:#FF0000;">*</font>:</label><br>
        <select name="igr_cod" id="igr_cod" class="selectpicker" data-size="8" title="Selecione..." data-live-search="true" onChange="mudaCampos(this.value)" parametro="no" requerido="yes" msg="Selecione a igreja!">	                    			           
        <%'LISTA ESTADOS
         SET QRY_IGR = conexao.execute("pr_le_igrejas 2, null")
         DO UNTIL QRY_IGR.EOF%>
         
            <option value="<%=QRY_IGR("igr_cod")%>" <%IF trim(QRY_IGR("igr_cod")) = trim(igr_cod) THEN response.write "selected"%>><%=QRY_IGR("igr_nFantasia")%> </option>
            
        <%QRY_IGR.MOVENEXT
        LOOP
        QRY_IGR.CLOSE			:	SET QRY_IGR 		= nothing%>           
		</select>
    </td>
    <td width="5"></td>
	<td>
        <label  class="control-label">NOME <font style="color:#FF0000;">*</font>:</label><br>
        <input type="text" name="irm_nome" id="irm_nome" value="<%=irm_nome%>" maxlength="70" class="form-control input-sm text-uppercase" title="Digite o nome!" parametro="no" requerido="yes" msg="Digite o nome do irmã(o)!" style="width:235px; height:30px">
    </td>
    <td width="5"></td>
    <td>
        <label  class="control-label">MÃE:</label><br>
        <input type="text" name="irm_nMae" id="irm_nMae" value="<%=irm_nMae%>" maxlength="70" class="form-control input-sm text-uppercase" title="Digite o nome da mãe!" parametro="no" requerido="no" msg="" style="width:200px; height:30px">
    </td>
    <td width="5"></td>
    <td>
        <label  class="control-label">PAI:</label><br>
        <input type="text" name="irm_nPai" id="irm_nPai" value="<%=irm_nPai%>" maxlength="70" class="form-control input-sm text-uppercase" title="Digite o nome da pai!" parametro="no" requerido="no" msg="" style="width:200px; height:30px">
    </td>  
    <td width="5"></td> 
    <td>
        <label class="control-label">CEP:</label><br>
        <input type="text" name="irm_cep" id="irm_cep" value="<%=irm_cep%>" maxlength="9" class="form-control input-sm text-uppercase" onKeyUp="mascaraCep(this); buscaCep(this.value)" onKeyDown="mascaraCep(this);" onKeypress="return somenteNumeros(event); mascaraCep(this); buscaCep(this.value);" onBlur="buscaCep(this.value)" placeholder="99999-999" parametro="no" requerido="no" msg="" style="width:130px;"> 
    </td>
</tr>
</table>

<br>
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>     	             
    <td>
        <label class="control-label">LOGRADOURO:</label><br>
        <input type="text" name="irm_logradouro" id="irm_logradouro" value="<%=irm_logradouro%>" maxlength="100" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:295px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">NÚMERO:</label><br>                                                
        <input type="text" name="irm_numero" id="irm_numero" value="<%=irm_numero%>" maxlength="8" class="form-control input-sm" onKeypress="return somenteNumeros(event);" placeholder="" parametro="no" requerido="no" msg="" style="width:80px;">
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">COMPLEMENTO:</label><br>                                                
        <input type="text" name="irm_complemento" id="irm_complemento" value="<%=irm_complemento%>" maxlength="20" class="form-control input-sm"  placeholder="" parametro="no" requerido="no" msg="" style="width:110px;">
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">BAIRRO:</label><br>
        <input type="text" name="irm_bairro" id="irm_bairro" value="<%=irm_bairro%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:150px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">CIDADE:</label><br>
        <input type="text" name="irm_cidade" id="irm_cidade" value="<%=irm_cidade%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:150px;">           
    </td>
    <td width="5"></td>
    
    <td>
       <label class="control-label">ESTADO:</label><br>
       <div class="input-group dropdown">
            <input type="text" name="irm_estado" id="irm_estado" onBlur="validaCampoOri()" class="form-control countrycode dropdown-toggle" maxlength="50" value="" readonly="readonly" parametro="no" requerido="no" msg="" style="width:160px;text-transform:uppercase">
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
    <td> 
        <label class="control-label">TELEFONE 1:</label><br>           
        <input type="text" name="irm_tel1" id="irm_tel1" value="<%=irm_tel1%>" onKeyUp="mascaraTel(this)" onKeyDown="mascaraTel(this)" onKeyPress="return somenteNumeros(event)" onBlur="verifNum(this.value, this.id)" maxlength="12" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:126px;">                            
    </td>
    <td width="5"></td>
    <td> 
        <label class="control-label">TELEFONE 2:</label><br>    
        <input type="text" name="irm_tel2" id="irm_tel2" value="<%=irm_tel2%>" onKeyUp="mascaraTel(this)" onKeyDown="mascaraTel(this)" onKeyPress="return somenteNumeros(event)" onBlur="verifNum(this.value, this.id)" maxlength="12" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:150px;">                            
    </td>
    <td width="5"></td>
    <td> 
        <label class="control-label">OPERADORA 1:</label><br>           
        <select name="irm_tpCel1" id="irm_tpCel1" class="selectpicker" data-size="8" title="Selecione..." data-live-search="true" parametro="no" requerido="no" msg="">	                    			
        <option value="">Selecione...</option>
	<option value="VIVO" <%IF TRIM(UCASE(irm_tpCel1)) = "VIVO" THEN RESPONSE.Write "SELECTED" END IF%>>VIVO</option>
        <option value="CLARO" <%IF TRIM(UCASE(irm_tpCel1)) = "CLARO" THEN RESPONSE.Write "SELECTED" END IF%>>CLARO</option>
        <option value="TIM" <%IF TRIM(UCASE(irm_tpCel1)) = "TIM" THEN RESPONSE.Write "SELECTED" END IF%>>TIM</option>
        <option value="OI" <%IF TRIM(UCASE(irm_tpCel1)) = "OI" THEN RESPONSE.Write "SELECTED" END IF%>>OI</option>
        <option value="NEXTEL" <%IF TRIM(UCASE(irm_tpCel1)) = "NEXTEL" THEN RESPONSE.Write "SELECTED" END IF%>>NEXTEL</option>
        <option value="ALGAR" <%IF TRIM(UCASE(irm_tpCel1)) = "ALGAR" THEN RESPONSE.Write "SELECTED" END IF%>>ALGAR</option>
        <option value="SERCOMTEL" <%IF TRIM(UCASE(irm_tpCel1)) = "SERCOMTEL" THEN RESPONSE.Write "SELECTED" END IF%>>SERCOMTEL</option>
        <option value="MVNO" <%IF TRIM(UCASE(irm_tpCel1)) = "MVNO" THEN RESPONSE.Write "SELECTED" END IF%>>MVNO (PORTO SEGURO, DATOR, TERAPAR</option>        
        <option value="OUTRO" <%IF TRIM(UCASE(irm_tpCel1)) = "OUTRO" THEN RESPONSE.Write "SELECTED" END IF%>>OUTRO (INTERNACIONAL)</option>
        </select>
    </td>
    <td width="5"></td>
    <td>                
        <label class="control-label">CELULAR 1:</label><br>
        <input type="text" name="irm_cel1" id="irm_cel1" value="<%=irm_cel1%>" onKeyUp="mascaraCel2(this)" onKeyDown="mascaraCel2(this)" onKeyPress="return somenteNumeros(event)" onBlur="verifNum(this.value, this.id)" maxlength="14" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:130px;">
    </td>
    <td width="5"></td>
    <td> 
        <label class="control-label">OPERADORA 2:</label><br>           
        <select name="irm_tpCel2" id="irm_tpCel2" class="selectpicker" data-size="8" title="Selecione..." data-live-search="true" parametro="no" requerido="no" msg="">	                    			
        <option value="">Selecione...</option>
	<option value="VIVO" <%IF TRIM(UCASE(irm_tpCel2)) = "VIVO" THEN RESPONSE.Write "SELECTED" END IF%>>VIVO</option>
        <option value="CLARO" <%IF TRIM(UCASE(irm_tpCel2)) = "CLARO" THEN RESPONSE.Write "SELECTED" END IF%>>CLARO</option>
        <option value="TIM" <%IF TRIM(UCASE(irm_tpCel2)) = "TIM" THEN RESPONSE.Write "SELECTED" END IF%>>TIM</option>
        <option value="OI" <%IF TRIM(UCASE(irm_tpCel2)) = "OI" THEN RESPONSE.Write "SELECTED" END IF%>>OI</option>
        <option value="NEXTEL" <%IF TRIM(UCASE(irm_tpCel2)) = "NEXTEL" THEN RESPONSE.Write "SELECTED" END IF%>>NEXTEL</option>
        <option value="ALGAR" <%IF TRIM(UCASE(irm_tpCel2)) = "ALGAR" THEN RESPONSE.Write "SELECTED" END IF%>>ALGAR</option>
        <option value="SERCOMTEL" <%IF TRIM(UCASE(irm_tpCel2)) = "SERCOMTEL" THEN RESPONSE.Write "SELECTED" END IF%>>SERCOMTEL</option>
        <option value="MVNO" <%IF TRIM(UCASE(irm_tpCel2)) = "MVNO" THEN RESPONSE.Write "SELECTED" END IF%>>MVNO (PORTO SEGURO, DATOR, TERAPAR</option>        
        <option value="OUTRO" <%IF TRIM(UCASE(irm_tpCel2)) = "OUTRO" THEN RESPONSE.Write "SELECTED" END IF%>>OUTRO (INTERNACIONAL)</option>
        </select>
    </td>
    <td width="5"></td>
    <td>                
        <label class="control-label">CELULAR 2:</label><br>
        <input type="text" name="irm_cel2" id="irm_cel2" value="<%=irm_cel2%>" onKeyUp="mascaraCel2(this)" onKeyDown="mascaraCel2(this)" onKeyPress="return somenteNumeros(event)" onBlur="verifNum(this.value, this.id)" maxlength="14" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:135px;">
    </td>       
</tr>
</table>

<br>
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>
	<td>
        <label class="control-label">NACIONALIDADE:</label><br>
        <input type="text" name="irm_nacionalidade" id="irm_nacionalidade" value="<%=irm_nacionalidade%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:190px;">           
    </td>
    <td width="5"></td>
    <td> 
        <label class="control-label">ESTADO CIVÍL:</label><br>           
        <select name="irm_estCivil" id="irm_estCivil" class="selectpicker" data-size="8" title="Selecione..." data-live-search="true" parametro="no" requerido="no" msg="">	                    			
        <option value="SOLTEIRO" <%IF TRIM(UCASE(irm_estCivil)) = "SOLTEIRO" THEN RESPONSE.Write "SELECTED" END IF%>>SOLTEIRO</option>
        <option value="CASADO" <%IF TRIM(UCASE(irm_estCivil)) = "CASADO" THEN RESPONSE.Write "SELECTED" END IF%>>CASADO</option>
        <option value="SEPARADO" <%IF TRIM(UCASE(irm_estCivil)) = "SEPARADO" THEN RESPONSE.Write "SELECTED" END IF%>>SEPARADO</option>
        <option value="DIVORCIADO" <%IF TRIM(UCASE(irm_estCivil)) = "DIVORCIADO" THEN RESPONSE.Write "SELECTED" END IF%>>DIVORCIADO</option>
        <option value="VIUVO" <%IF TRIM(UCASE(irm_estCivil)) = "VIUVO" THEN RESPONSE.Write "SELECTED" END IF%>>VIÚVO</option>
        <option value="OUTRO" <%IF TRIM(UCASE(irm_estCivil)) = "OUTRO" THEN RESPONSE.Write "SELECTED" END IF%>>OUTRO</option>
        </select>
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">NOME CONJUGUE:</label><br>
        <input type="text" name="irm_nConjegue" id="irm_nConjegue" value="<%=irm_nConjegue%>" maxlength="70" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:200px;">           
    </td>
    <td width="5"></td>
    <td> 
        <label class="control-label">NÚM FILHOS:</label><br>           
        <input type="text" name="irm_qtdFilhos" id="irm_qtdFilhos" value="<%=irm_qtdFilhos%>" onKeyPress="return somenteNumeros(event)" maxlength="2" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:100px;">                            
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">RG:</label><br>
        <input type="text" name="irm_rg" id="irm_rg" value="<%=irm_rg%>" maxlength="20" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:100px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">CPF/ID (INTERNACIONAL):</label><br>
        <input type="text" name="irm_cpf" id="irm_cpf" value="<%=irm_cpf%>" maxlength="20" class="form-control input-sm" parametro="no" requerido="no" msg="" style="width:170px;">           
    </td>
</tr>
</table>

<br>
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>
	<td> 
        <label class="control-label">CRENTE?</label><br>           
        <select name="irm_crente" id="irm_crente" class="selectpicker" data-size="4" title="Selecione..." data-live-search="true" parametro="no" requerido="no" msg="">	                    			
        <option value="S" <%IF TRIM(UCASE(irm_crente)) = "S" THEN RESPONSE.Write "SELECTED" END IF%>>SIM</option>
        <option value="N" <%IF TRIM(UCASE(irm_crente)) = "N" THEN RESPONSE.Write "SELECTED" END IF%>>NÃO</option>
        </select>
    </td>
    <td width="5"></td>
    <td> 
        <label class="control-label">BATIZADO NO ESPÍRITO?</label><br>           
        <select name="irm_btEspiritos" id="irm_btEspiritos" class="selectpicker" data-size="4" title="Selecione..." data-live-search="true" parametro="no" requerido="no" msg="">	                    			
        <option value="S" <%IF TRIM(UCASE(irm_btEspiritos)) = "S" THEN RESPONSE.Write "SELECTED" END IF%>>SIM</option>
        <option value="N" <%IF TRIM(UCASE(irm_btEspiritos)) = "N" THEN RESPONSE.Write "SELECTED" END IF%>>NÃO</option>
        </select>
    </td>

    <!--<td width="5"></td>
    <td>
        <label class="control-label">DATA DO BATISMO:</label><br>            
        <input type="text" name="irm_dtBtEspiritos" id="irm_dtBtEspiritos" value="<%=irm_dtBtEspiritos%>" onKeyPress="return somenteNumeros(event);barra(this)" onKeyUp="barra(this)" onKeyDown="barra(this)" onBlur="verifNum(this.value, this.id)" maxlength="10" class="form-control input-sm" parametro="data" requerido="no" msg="" style="width:160px;">
    </td>-->


    <td width="5"></td>
    <td> 
        <label class="control-label">BATIZADO NAS ÁGUAS?</label><br>           
        <select name="irm_btAguas" id="irm_btAguas" class="selectpicker" data-size="4" title="Selecione..." data-live-search="true" parametro="no" requerido="no" msg="">	                    			
        <option value="S" <%IF TRIM(UCASE(irm_btAguas)) = "S" THEN RESPONSE.Write "SELECTED" END IF%>>SIM</option>
        <option value="N" <%IF TRIM(UCASE(irm_btAguas)) = "N" THEN RESPONSE.Write "SELECTED" END IF%>>NÃO</option>
        </select>
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">DATA DO BATISMO:</label><br>            
        <input type="text" name="irm_dtBtAguas" id="irm_dtBtAguas" value="<%=irm_dtBtAguas%>" onKeyPress="return somenteNumeros(event);barra(this)" onKeyUp="barra(this)" onKeyDown="barra(this)" onBlur="verifNum(this.value, this.id)" maxlength="10" class="form-control input-sm" parametro="data" requerido="no" msg="" style="width:330px;">
    </td>
</tr>
</table>

<br>
<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
<tr>
	<td>
        <label class="control-label">DENOMINAÇÃO DE BATISMO:</label><br>            
        <input type="text" name="irm_denominacaoBatismo" id="irm_denominacaoBatismo" value="<%=irm_denominacaoBatismo%>" maxlength="70" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:200px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">CIDADE / ESTADO DOS PAIS:</label><br>            
        <input type="text" name="irm_CidEstPais" id="irm_CidEstPais" value="<%=irm_CidEstPais%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:230px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">DATA DO INGRESSO:</label><br>            
        <input type="text" name="irm_dtIngresso" id="irm_dtIngresso" value="<%=irm_dtIngresso%>" onKeyPress="return somenteNumeros(event);barra(this)" onKeyUp="barra(this)" onKeyDown="barra(this)" onBlur="verifNum(this.value, this.id)" maxlength="10" class="form-control input-sm" parametro="data" requerido="no" msg="" style="width:200px;">
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">FUNÇÃO NA IGREJA:</label><br>            
        <input type="text" name="irm_funcIgreja" id="irm_funcIgreja" value="<%=irm_funcIgreja%>" maxlength="50" class="form-control input-sm text-uppercase" parametro="no" requerido="no" msg="" style="width:170px;">           
    </td>
    <td width="5"></td>
    <td>
        <label class="control-label">Dt Nascimento do Irmão:</label><br>            
        <input type="text" name="irm_dtNascimento" id="irm_dtNascimento" value="<%=irm_dtNascimento%>" onKeyPress="return somenteNumeros(event)" onKeyUp="barra(this)" onBlur="verifNum(this.value, this.id)" maxlength="10" class="form-control input-sm" parametro="data" requerido="no" msg="" style="width:180px;">           
    </td>
</tr>
</table>

<table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px; width:85%">
<tr>
	<td>
        <!--Textarea with icon prefix-->
        <div class="md-form mb-4 blue-textarea active-blue-textarea">
          <br>
          <textarea id="irm_obs" name="irm_obs" class="md-textarea form-control" rows="5" maxlength="10000"><%=irm_obs%></textarea>
          <label for="form21" style="margin-top:5px">QUAISQUER OBSERVAÇÕES E/OU PARTICULARIDADES DOS IRMÃOS</label>
        </div>       
        
    </td>
</tr>
</table>

<br>
<table cellpadding="2" cellspacing="0" border="0" class="formulario" align="center">
<tr>
    <td>
		<%IF LEN(irm_cod) =0 THEN%>
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

</body>

<!-- MODAL LISTA IRMAOS-->
<div class="modal fade form-horizontal"  id="irmaosModal" tabindex="-1" role="dialog"  aria-labelledby="myModalLabel" aria-hidden="true">
    <div class="modal-dialog" style="width: 850px;" role="document">
        <div class="modal-content" style="top:-30px !important">
            <!-- Modal Header -->
            <div class="modal-header" style="background:#003d7b">                
                <h4 class="modal-title" id="myModalLabel" style="color:#FFFFFF;font-weight:bold;font-size:16px">
                    LISTA IRMÃOS
                </h4>
            </div>
            
            <!-- Modal Body -->
            <div class="modal-body">
        		<iframe name="iframeListaIrmaos" id="iframeListaIrmaos" src="listaIrm.asp" width="100%" height="70%" frameborder="0" allowfullscreen="allowFullScreen" scrolling="no"></iframe>           
            </div>
            
            <!-- Modal Footer -->
            <div class="modal-footer"> 
				<input type="button" id="fecharirmaosModal" name="fecharirmaosModal" value="Fechar" class="btn btn-primary" style="background:#003d7b;">
            </div>    

        </div>
    </div>   
    
</div>

<script>

// faz função do select na lista nao ordenada estado
$(function() {
  $('#estado a').click(function() {
	console.log($(this).attr('data-value'));
	$(this).closest('.dropdown').find('input.countrycode')
	  .val($(this).attr('data-value'))	
  });
});

// funcao de acao do botao
$('#fecharirmaosModal').on('click', function() {
	$("#irmaosModal").modal('toggle');	
	// desbloqueia menu
	parent.desblockMenu()
});

// verifica se é internacional
function mudaCampos(valor){
	
	//busca igr
	frameaux.location = 'ajax/iframe_geral.asp?tipo=3&cod=' + valor	
}

//ajax para buscar cep nacional
function buscaCep(cep){
	
	if(cep.length == 9){	 
		frameaux.location = 'ajax/iframe_geral.asp?tipo=4&cep=' + cep
	}	 
}

<%IF wstatus = "SIM" THEN
	'VERIFICA SE É INTERNACIONAL
	IF igr_internacional = "S" THEN%>
		// coloca valor no campo			
		$('#irm_estado').val('<%=irm_estado%>')		
		
		// chama funcao
		formatCampos(1)
	<%ELSE%>
	
		// coloca valor no campo			
		$('#irm_estado').val('<%=irm_estado%>')		
		
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
		
		// formata cep
		$("#irm_cep").attr('maxlength' , '15')
		$("#irm_cep").attr('onKeyUp' , '')
		$("#irm_cep").attr('onKeyDown' , '')
		$("#irm_cep").attr('onKeypress' , '')
		$("#irm_cep").attr('onBlur' , '')
		$("#irm_cep").attr('placeholder' , '###########')
		
		// formata estado
		$('#irm_estado').prop('readonly', false);
		
		// formata campo telefones
		$("#irm_tel1").attr('maxlength' , '20')
		$("#irm_tel2").attr('maxlength' , '20')
		$("#irm_cel1").attr('maxlength' , '20')
		$("#irm_cel2").attr('maxlength' , '20')
		
		$("#irm_tel1").attr('onKeyUp' , '')
		$("#irm_tel1").attr('onKeyDown' , '')		
		$("#irm_tel1").attr('onKeyPress' , '')
		$("#irm_tel1").attr('onBlur' , '')
		
		$("#irm_tel2").attr('onKeyUp' , '')
		$("#irm_tel2").attr('onKeyDown' , '')
		$("#irm_tel2").attr('onKeyPress' , '')
		$("#irm_tel2").attr('onBlur' , '')
		
		$("#irm_cel1").attr('onKeyUp' , '')
		$("#irm_cel1").attr('onKeyDown' , '')		
		$("#irm_cel1").attr('onKeyPress' , '')
		$("#irm_cel1").attr('onBlur' , '')
		
		$("#irm_cel2").attr('onKeyUp' , '')		
		$("#irm_cel2").attr('onKeyDown' , '')		
		$("#irm_cel2").attr('onKeyPress' , '')
		$("#irm_cel2").attr('onBlur' , '')
		
		$("#irm_tel1").attr('placeholder' , '############')
		$("#irm_tel2").attr('placeholder' , '############')
		$("#irm_cel1").attr('placeholder' , '############')
		$("#irm_cel2").attr('placeholder' , '############')
	}
	
	if(tipo == 2){		
	
		// formata cep
		$("#irm_cep").attr('maxlength' , '9')
		$("#irm_cep").attr('onKeyUp' , 'mascaraCep(this); buscaCep(this.value)')
		$("#irm_cep").attr('onKeyDown' , 'mascaraCep(this)')
		$("#irm_cep").attr('onKeypress' , 'return somenteNumeros(event); mascaraCep(this); buscaCep(this.value);')
		$("#irm_cep").attr('onBlur' , 'buscaCep(this.value)')
		$("#irm_cep").attr('placeholder' , '99999-999')	
		
		// formata estado
		$('#irm_estado').prop('readonly', true);
		
		// formata campo telefones
		$("#irm_tel1").attr('maxlength' , '12')
		$("#irm_tel2").attr('maxlength' , '12')
		$("#irm_cel1").attr('maxlength' , '13')
		$("#irm_cel2").attr('maxlength' , '13')
		
		$("#irm_tel1").attr('onKeyUp' , 'mascaraTel(this)')
		$("#irm_tel1").attr('onKeyDown' , 'mascaraTel(this)')
		$("#irm_tel1").attr('onKeyPress' , 'return somenteNumeros(event)')
		$("#irm_tel1").attr('onBlur' , 'verifNum(this.value, this.id)')
		
		$("#irm_tel2").attr('onKeyUp' , 'mascaraTel(this)')
		$("#irm_tel2").attr('onKeyDown' , 'mascaraTel(this)')
		$("#irm_tel2").attr('onKeyPress' , 'return somenteNumeros(event)')
		$("#irm_tel2").attr('onBlur' , 'verifNum(this.value, this.id)')
		
		$("#irm_cel1").attr('onKeyUp' , 'mascaraCel2(this)')
		$("#irm_cel1").attr('onKeyDown' , 'mascaraCel2(this)')
		$("#irm_cel1").attr('onKeyPress' , 'return somenteNumeros(event)')
		$("#irm_cel1").attr('onBlur' , 'verifNum(this.value, this.id)')
		
		$("#irm_cel2").attr('onKeyUp' , 'mascaraCel2(this)')
		$("#irm_cel2").attr('onKeyDown' , 'mascaraCel2(this)')
		$("#irm_cel2").attr('onKeyPress' , 'return somenteNumeros(event)')
		$("#irm_cel2").attr('onBlur' , 'verifNum(this.value, this.id)')
		
		$("#irm_tel1").attr('placeholder' , '99 9999-9999')
		$("#irm_tel2").attr('placeholder' , '99 9999-9999')
		$("#irm_cel1").attr('placeholder' , '99 99999-9999')
		$("#irm_cel2").attr('placeholder' , '99 99999-9999')
	}
	
}

function abreModal(){	
	
	$('#irmaosModal').modal({backdrop: 'static', keyboard: true}); 
	$('#irmaosModal').modal('show');	
	
	// bloqueia menu
	parent.blockMenu()
}
</script>

<script>
////////////////////////////////////// erro
<%IF erro = 0 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Irmã(o) cadastrado com sucesso!')
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
	$('#avisoDescricao').text('Irmã(o) já cadastrado!')
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
	$('#avisoDescricao').text('Erro ao cadastrar irmã(o), contate o suporte!')
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
	$('#avisoDescricao').text('Irmã(o) alterado com sucesso!')
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
	$('#avisoDescricao').text('Irmã(o) excluído com sucesso!')
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