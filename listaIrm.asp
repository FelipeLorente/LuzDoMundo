<!--#include file="include/conexao.asp"-->
<!--#include file="include/topoPaginacao.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<%'PARAMETROS OBRIGATORIOS
wprograma 		= "listaIrm"
igr_cod			= request.Form("igr_cod")
irm_nome		= request.Form("irm_nome")
irm_funcIgreja	= request.Form("irm_funcIgreja")

'VERIFICA A SESSION
IF LEN(SESSION("usu_cod")) = 0 THEN
	REDIRECTPAGE(3)	
	response.End() 
END IF

' OBS aaaaaaaaaaaaaaaaaaaa
'coloca o caminho da atualiza��o
URL = wprograma&".asp"
		
'RECUPERA VALORES DO BANCO
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conexao
cmd.CommandText = "pr_le_irmaosBusca"
cmd.CommandType = 4
Set params = cmd.Parameters

params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
params.Append cmd.CreateParameter("@igr_cod", 3, 1, 0, numerosql(igr_cod))
params.Append cmd.CreateParameter("@irm_nome", 200, 1, 70, stringsqlserver(irm_nome))
params.Append cmd.CreateParameter("@irm_funcIgreja", 200, 1, 50, stringsqlserver(irm_funcIgreja))		

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
END IF%>

<div id="mostraForm" style="width:98% !important; margin-left:15px;">    
     
    <form name="forme" id="forme" action="<%=wprograma%>.asp" method="post" class="form-inline" role="form">	
    <table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
    <tr>   
        <td>
            <label class="control-label">IGREJA:</label><br>
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
        <td width="10"></td>
        <td>
            <label class="control-label">IRM�(O):</label><br>
            <input type="text" name="irm_nome" id="irm_nome" value="<%=irm_nome%>" maxlength="70" class="form-control text-uppercase" title="Digite do nome do irm�(o)!" parametro="no" requerido="no" msg=""  style="width:250px">
        </td>
        <td width="10"></td> 
        <td>
            <label class="control-label">FUN��O NA IGREJA:</label><br>
            <input type="text" name="irm_funcIgreja" id="irm_funcIgreja" value="<%=irm_funcIgreja%>" maxlength="50" class="form-control text-uppercase" title="Digite a fun��o..!" parametro="no" requerido="no" msg=""  style="width:200px">
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
                    <th scope="col" title="Baixar PDF?">PDF</th>                    
                    <th scope="col">IRM�(O)</th>                    
                    <th scope="col">TELEFONE</th>
                    <th scope="col">CELULAR</th> 
                    <th scope="col" title="Exclu�r este irm�(o)!">EXCLU�R?</th>                                                      
	            </tr>              
            </thead>
            
            <%'LOOP 
			DO WHILE count < RegPorPag AND NOT QRY.EOF
			count		= count + 1	
				
				'RECUPERA CAMPOS
				igr_cod					= QRY("igr_cod") 
				igr_cnpj 				= QRY("igr_cnpj")
				igr_rSocial 			= QRY("igr_rSocial") 
				irm_cod					= QRY("irm_cod") 
				irm_nome				= QRY("irm_nome") 
				irm_tel1				= QRY("irm_tel1") 
				irm_tpCel1				= QRY("irm_tpCel1")
				irm_cel1				= QRY("irm_cel1") 
				irm_funcIgreja			= QRY("irm_funcIgreja") 
				
				IF LEN(TRIM(irm_tpCel1)) >0 THEN
					CELULAR = irm_tpCel1 & " | " & irm_cel1
				ELSE
					CELULAR = irm_cel1	
				END IF
				
				'MONTA ENDERE�O
				endereco = "ENDERE�O: " & igr_logradouro & ", " & igr_numero & " BAIRRO: " & igr_bairro & ", CIDADE: " & igr_cidade & " / " & igr_estado & " (" & retornastatus(igr_status) & ")"
				
				'MONTA LOCATION
				locationIrm = "irmaos.asp?irm_cod="&irm_cod&""
				
				locationExc = "irmaos.asp?tipo=3&irm_cod="&irm_cod&""%>
            	
                <tbody>
                    <tr title="<%="Alterar? | " & UCASE(irm_nome) & " membro da igreja: " & igr_rSocial & "| Fun��o: "&irm_funcIgreja%>" style="cursor:pointer">
	                    <td title="Baixar PDF?" onClick="abreJanela('<%=irm_cod%>')"><img src="imagens/pdf.png" width="20" height="20"></td>
                        <td onClick="alteraReg('<%=locationIrm%>')"><%=irm_nome%></td>
                        <td onClick="alteraReg('<%=locationIrm%>')"><%=irm_tel1%></td>
                        <td onClick="alteraReg('<%=locationIrm%>')"><%=CELULAR%></td>
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
        
	<%END IF%>
</div>

</body>
</html>

<!--#include file="include/bottom.asp"-->

<script>
function deletaIgr(links){
	
	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#avisoConfirma').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#avisoConfirma').modal('show')
	$('#avisoTituloConfirma').text('Aviso do Sistema')
	$('#avisoDescricaoConfirma').text('Deseja excluir este irm�(o)?')
	$('#avisoImagemConfirma').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	$('#avisoConfirma').ready(function(e) {
        $('#nao').focus();	
    });				
    
	//foca o campo quando apertar ok
	$("#sim").click(function(){
		// desbloqueia menu
		parent.parent.desblockMenu()									
		parent.location=links
	});
}

function alteraReg(links){
	
	// desbloqueia menu
	parent.parent.desblockMenu()
	
	parent.location = links
}

function abreJanela(cod){
	
	// abre janela para salvar em PDF
	window.open('montaPDF.asp?abc=' + cod, '_blank');

}
</script>

<%'FECHA TODAS AS QRY'S E CONEX�ES COM O BANCO DE DADOS
conexao.close			:	SET conexao 		= nothing%>