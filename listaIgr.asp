<!--#include file="include/conexao.asp"-->
<!--#include file="include/topoPaginacao.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<%'PARAMETROS OBRIGATORIOS
wprograma 			= "listaIgr"
igr_cnpj			= request.Form("igr_cnpj")
igr_nFantasia			= request.Form("igr_nFantasia")
igr_responsavel		= request.Form("igr_responsavel")

'VERIF SE TEM CNPJ
IF LEN(TRIM(igr_cnpj)) >0 THEN
	igr_cnpj = retiraCaracteres(igr_cnpj)
END IF

'VERIFICA A SESSION
IF LEN(SESSION("usu_cod")) = 0 THEN
	REDIRECTPAGE(3)	
	response.End() 
END IF

' OBS
'coloca o caminho da atualiza��o
URL = wprograma&".asp"
		
'RECUPERA VALORES DO BANCO
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conexao
cmd.CommandText = "pr_le_igrejasBusca"
cmd.CommandType = 4
Set params = cmd.Parameters

params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
params.Append cmd.CreateParameter("@igr_cnpj", 200, 1, 14, stringsqlserver(igr_cnpj))
params.Append cmd.CreateParameter("@igr_nFantasia", 200, 1, 60, stringsqlserver(igr_nFantasia))
params.Append cmd.CreateParameter("@igr_responsavel", 200, 1, 50, stringsqlserver(igr_responsavel))		

SET QRY = cmd.execute	

'PAGINA��O
SET QRY = Server.CreateObject("ADODB.RecordSet")
	QRY.CursorType 		= 3
	QRY.CursorLocation 	= 3						'exibe dados atualizados 
	QRY.lockType		= 3 					'define cursor com dados est�ticos
	QRY.cursorLocation 	= 3 					'cursor no cliente
	QRY.open cmd	

'VERIFICA SE � FINAL DE ARQUIVO
IF QRY.EOF THEN
	wstatus				    = "no"
ELSE
	wstatus 				= "yes"
	QRY.PageSize			= RegPorPag 		'Numero de registros por p�gina
	registros				= QRY.RecordCount
	total					= QRY.PageSize
	paginas					= QRY.pageCount
	registro_atual			= QRY.absolutePosition
	QRY.AbsolutePage 		= PagAtual
	TotPag 					= QRY.PageCount
END IF

'FORMATA CNPJ
IF LEN(TRIM(igr_cnpj)) = 14 THEN
	igr_cnpj = FormataCNPJ(igr_cnpj)
ELSE
	igr_cnpj = ""
END IF%>

<style>
	body{ overflow:hidden;}
</style>

<div id="mostraForm" style="width:98% !important; margin-left:15px;">    
     
    <form name="forme" id="forme" action="<%=wprograma%>.asp" method="post" class="form-inline" role="form">	
    <table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
    <tr>   
        <!--<td>
            <label class="control-label">CNPJ:</label><br>
            <input type="text" name="igr_cnpj" id="igr_cnpj" value="<%=igr_cnpj%>" maxlength="18" onKeyUp="mascaraCnpj(this)" onKeyDown="mascaraCnpj(this)" onKeypress="return somenteNumeros(event);" class="form-control" placeholder="CNPJ: 99.999.999/9999-99" title="Digite o cnpj!" style="width:200px">
        </td>
        <td width="10"></td>-->
        <td>
            <label class="control-label">IGREJA:</label><br>
            <input type="text" name="igr_nFantasia" id="igr_nFantasia" value="<%=igr_nFantasia%>" maxlength="60" class="form-control text-uppercase" placeholder="IGREJA 1.." title="Digite do nome da igreja!" parametro="no" requerido="no" msg=""  style="width:250px">
        </td>
        <td width="10"></td> 
        <td>
            <label class="control-label">RESPONS�VEL:</label><br>
            <input type="text" name="igr_responsavel" id="igr_responsavel" value="<%=igr_responsavel%>" maxlength="50" class="form-control text-uppercase" placeholder="RESPONS�VEL.." title="Digite o respons�vel!" parametro="no" requerido="no" msg=""  style="width:200px">
        </td>
        <td width="10"></td>        
        <td>
            <br>	
            <button type="submit" id="buscar" name="buscar" value="1" style="background:#003d7b;font-weight:bold;font-size:15px" class="btn btn-md btn-primary btn-block"> BUSCAR</button> </td>
    </tr>
    </table>                                                           
    </form>
                
        
	 <%IF wstatus = "yes" THEN%>       
        <div class="table-responsive-sm">    		                       
        
           <table class="table table-striped" style="font-size:12px">
           <thead>
           		<tr>
                    <th scope="col">CNPJ</th>                    
                    <th scope="col">IGREJA</th>
                    <th scope="col">RESPONS�VEL</th>                    
                    <th scope="col">CELULAR</th>
                    <th scope="col" title="Exclu�r este arquivos!">EXCLU�R?</th>                                                      
	            </tr>              
            </thead>
            
            <%'LOOP 			
			DO WHILE count < RegPorPag AND NOT QRY.EOF
			count		= count + 1	
				
				'RECUPERA CAMPOS
				igr_cod					= QRY("igr_cod") 
				igr_cnpj 				= QRY("igr_cnpj")
				igr_rSocial 			= QRY("igr_rSocial") 
				igr_nFantasia 			= QRY("igr_nFantasia") 
				igr_cep 				= QRY("igr_cep") 
				igr_logradouro 			= QRY("igr_logradouro") 
				igr_numero 				= QRY("igr_logradouro") 
				igr_complemento 		= QRY("igr_complemento") 
				igr_bairro 				= QRY("igr_bairro")
				igr_cidade 				= QRY("igr_cidade") 
				igr_estado 				= QRY("igr_estado") 
				igr_status 				= QRY("igr_status") 
				igr_responsavel 		= QRY("igr_responsavel") 
				igr_telResponsalvel 	= QRY("igr_telResponsalvel") 
				
				'MONTA ENDERE�O
				endereco = "ENDERE�O: " & igr_logradouro & ", " & igr_numero & " BAIRRO: " & igr_bairro & ", CIDADE: " & igr_cidade & " / " & igr_estado & " (" & retornastatus(igr_status) & ")"
				
				'MONTA LOCATION
				locationIgr = "parent.location='igrejas.asp?altera=sim&igr_cod="&igr_cod&"'"
				
				locationExc = "igrejas.asp?altera=sim&tipo=3&igr_cod="&igr_cod&""%>
            	
                <tbody>
                    <tr title="<%="Alterar? | " & UCASE(endereco)%>" style="cursor:pointer">
                        <td onClick="<%=locationIgr%>"><%=FormataCNPJ(igr_cnpj)%></td>
                        <td onClick="<%=locationIgr%>"><%=igr_nFantasia%></td>
                        <td onClick="<%=locationIgr%>"><%=igr_responsavel%></td>
                        <td onClick="<%=locationIgr%>"><%=igr_telResponsalvel%></td>
                        <td title="Deseja exclu�r esta igreja?" style="font-size:14px; font-weight:bold; color:red; padding-left:30px">
                        	<a onClick="deletaIgr('<%=locationExc%>')" style="text-decoration:none; color:red">X</a>
                        </td>
                    </tr>
				</tbody>
            
            <%QRY.MOVENEXT
            LOOP
            QRY.CLOSE			:	SET QRY 		= nothing%>             	        
        
    	</table> 
        
        <table cellpadding="0" cellspacing="0" align="center">
        <tr>
            <td colspan="99"><%LinksNavegacao()%> </td>
        </tr>
        </table>
    
    </div>

</div>

<%END IF%>

</body>
</html>

<!--#include file="include/bottom.asp"-->

<script>
function deletaIgr(links){
	
	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	parent.$('#avisoConfirma').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	parent.$('#avisoConfirma').modal('show')
	parent.$('#avisoTituloConfirma').text('Aviso do Sistema')
	parent.$('#avisoDescricaoConfirma').text('Ao exclu�r esta igreja todos os IRM�OS vinculados a ela ser�o exclu�dos tamb�m. Deseja prosseguir com a exclus�o?')
	parent.$('#avisoImagemConfirma').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	parent.$('#avisoConfirma').ready(function(e) {
        parent.$('#nao').focus();	
    });				
    
	//foca o campo quando apertar ok
	parent.$("#sim").click(function(){
		parent.location=links
	});
}
</script>

<%'FECHA TODAS AS QRY'S E CONEX�ES COM O BANCO DE DADOS
conexao.close			:	SET conexao 		= nothing%>