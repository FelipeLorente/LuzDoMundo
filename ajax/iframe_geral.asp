<!--#include file="../include/conexao.asp"-->

<%'PARÂMETROS OBRIGATÓRIOS
wprograma		= "iframe_geral"
wvariavel		= null
tipo			= request.querystring("tipo")
	
IF tipo = 1 THEN
	
    'RECUPERA PARAMETROS
    cep_cep = Request.Querystring("cep")
    
    'RETIRA CARACTERES
	 cep_cep = RETIRACARACTERES(cep_cep)
    
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_le_cep"
	cmd.CommandType = 4
	Set params = cmd.Parameters

	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
	params.Append cmd.CreateParameter("@cep_cep", 200, 1, 10, cep_cep)

	set QRY = cmd.Execute

	IF NOT QRY.EOF THEN%>
		<script>	
			
			//coloca valores nos campos
			parent.$('#igr_numero').val('')
			parent.$('#igr_complemento').val('')
			parent.$('#igr_logradouro').val('<%=QRY("cep_tpLogradouro") & " " & QRY("cep_endereco")%>')
			parent.$('#igr_bairro').val('<%=QRY("cep_bairro")%>')
			parent.$('#igr_cidade').val('<%=QRY("cep_cidade")%>')
			parent.$("#igr_estado").val("<%=QRY("cep_estadoSigla") & " - " & QRY("cep_estado")%>").change();
			parent.$('#igr_numero').focus()
				
		</script>

	<%	QRY.close			:	SET QRY 		= nothing
	ELSE%>
		<script>		
			parent.$('#igr_logradouro').val('')
			parent.$('#igr_bairro').val('')
			parent.$('#igr_cidade').val('')
			
			//FOCA CAMPO
			parent.$('#igr_logradouro').focus()		
		</script>    
	<%END IF

END IF


IF tipo = 2 THEN
	
	'RECUPERA PARAMETROS
    usu_cod = Request.Querystring("usu_cod")
	
	'EXCLUI REGISTRO
	SET QRY = conexao.Execute("pr_ma_usuarios 3, " &usu_cod& ", NULL, NULL, NULL, NULL, NULL")
	
	'VERIF SE TEM RETORNO
	IF NOT QRY.EOF THEN
	
		MSG = QRY("MSG")
		
		QRY.close			:	SET QRY 		= nothing%>
		
        <script>
        
		<%IF MSG = 1 THEN%>
			
			// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
			parent.$('#aviso').modal({backdrop: 'static', keyboard: false}) 
			
			// mostra aviso de erro
			parent.$('#aviso').modal('show')
			parent.$('#avisoTitulo').text('Aviso do Sistema')
			parent.$('#avisoDescricao').text('Usuário excluído com sucesso!')
			parent.$('#avisoImagem').addClass('glyphicon-ok text-success');
			
			// foco no botao
			parent.$('#aviso').ready(function(e) {
				parent. $('#ok').focus();	
			});				
			
			//foca o campo quando apertar ok
			parent.$("#ok").click(function(){
				parent.location.href = 'listaUsu.asp'													
			});
			
		<%ELSE%>
		
			// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
			parent.$('#aviso').modal({backdrop: 'static', keyboard: false}) 
			
			// mostra aviso de erro
			parent.$('#aviso').modal('show')
			parent.$('#avisoTitulo').text('Aviso do Sistema')
			parent.$('#avisoDescricao').text('usuário não excluído!')
			parent.$('#avisoImagem').addClass('glyphicon-remove text-danger');
			
			// foco no botao
			parent.$('#aviso').ready(function(e) {
				 parent.$('#ok').focus();	
			});				
			
			//foca o campo quando apertar ok
			parent.$("#ok").click(function(){
				parent.location.href = 'listaUsu.asp'													
			});		
		<%END IF%>
		
		</script>
	<%END IF

END IF

IF tipo = 3 THEN

	'RECUPERA CAMPOS
	igr_cod = Request.QueryString("cod")		
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_le_igrejas"
	cmd.CommandType = 4
	Set params = cmd.Parameters

	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
	params.Append cmd.CreateParameter("@tipo", 3, 1, 0, 1)
	params.Append cmd.CreateParameter("@igr_cod", 3, 1, 0, igr_cod)

	set QRY = cmd.Execute%>

	<script>

	<%IF NOT QRY.EOF THEN
		
		'RECUPERA CAMPOS
		igr_internacional = QRY("igr_internacional")
		
		QRY.close			:	SET QRY 		= nothing	%>		        
		
		// limpa campos do endereco
		parent.$('#irm_cep').val('')
		parent.$('#irm_logradouro').val('')
		parent.$('#irm_numero').val('')
		parent.$('#irm_complemento').val('')
		parent.$('#irm_bairro').val('')
		parent.$('#irm_cidade').val('')
		parent.$('#irm_estado').val('')								
		
		<%'VERIF SE E INTERNACIONAL
		IF UCASE(TRIM(igr_internacional)) = "S" THEN%>
			
			parent.$('#irm_estado').prop('readonly', false);
			
			// chama funcao
			parent.formatCampos(1)            

        <%ELSE%>
		
			parent.$('#irm_estado').prop('readonly', true);
			
			// chama funcao
			parent.formatCampos(2)            		
                    	    
	<%	END IF		
	ELSE%>

		parent.$('#irm_estado').prop('readonly', true);
		
		// chama funcao
		parent.formatCampos(2)					

	<%END IF%>
	
	</script>
	
<%END IF

IF tipo = 4 THEN
	
    'RECUPERA PARAMETROS
    cep_cep = Request.Querystring("cep")
    
    'RETIRA CARACTERES
	 cep_cep = RETIRACARACTERES(cep_cep)
    
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conexao
	cmd.CommandText = "pr_le_cep"
	cmd.CommandType = 4
	Set params = cmd.Parameters

	params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
	params.Append cmd.CreateParameter("@cep_cep", 200, 1, 10, cep_cep)

	set QRY = cmd.Execute

	IF NOT QRY.EOF THEN%>
		<script>	
			
			//coloca valores nos campos
			parent.$('#irm_numero').val('')
			parent.$('#irm_complemento').val('')
			parent.$('#irm_logradouro').val('<%=QRY("cep_tpLogradouro") & " " & QRY("cep_endereco")%>')
			parent.$('#irm_bairro').val('<%=QRY("cep_bairro")%>')
			parent.$('#irm_cidade').val('<%=QRY("cep_cidade")%>')
			parent.$("#irm_estado").val("<%=QRY("cep_estadoSigla")%>").change();
			parent.$('#irm_numero').focus()
				
		</script>

	<%	QRY.close			:	SET QRY 		= nothing
	ELSE%>
		<script>		
			parent.$('#irm_logradouro').val('')
			parent.$('#irm_bairro').val('')
			parent.$('#irm_cidade').val('')
			
			//FOCA CAMPO
			parent.$('#irm_logradouro').focus()		
		</script>    
	<%END IF

END IF%>    

<!--#include file="../include/fechaConexao.asp"-->