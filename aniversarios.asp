<!--#include file="include/conexao.asp"-->
<!--#include file="include/topoPaginacaoUsuario.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<%'PARAMETROS OBRIGATORIOS
wprograma 				= "aniversarios"
irm_nome  				= request.Form("irm_nome")
irm_dtNascimento		= request.Form("irm_dtNascimento")
tipo					= request.QueryString("tipo")

'VERIFICA A SESSION
IF LEN(SESSION("usu_cod")) = 0 THEN
	REDIRECTPAGE(3)	
	response.End() 
END IF

IF tipo = 1 THEN
	'RECUPERA CAMPOS
	irm_cod = request.QueryString("irm_cod")
	
	IF LEN(TRIM(irm_cod))>0 THEN
		conexao.execute("pr_ma_baixaNiver 1 , "&irm_cod&"")
		
		WERRO = 1
	END IF
END IF

IF tipo = 2 THEN
		
	conexao.execute("pr_ma_baixaNiver 2 , null")
		
	WERRO = 2	
END IF

' OBS
'coloca o caminho da atualização
URL = wprograma&".asp"
		
'RECUPERA VALORES DO BANCO
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conexao
cmd.CommandText = "pr_le_aniversariosBusca"
cmd.CommandType = 4
Set params = cmd.Parameters

params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
params.Append cmd.CreateParameter("@irm_nome", 200, 1, 70, stringsqlserver(irm_nome))
params.Append cmd.CreateParameter("@irm_dtNascimento", 200, 1, 50, stringsqlserver(irm_dtNascimento))

SET QRY = cmd.execute	

'PAGINAÇÃO
SET QRY = Server.CreateObject("ADODB.RecordSet")
	QRY.CursorType 		= 3
	QRY.CursorLocation 	= 3						'exibe dados atualizados 
	QRY.lockType		= 3 					'define cursor com dados estáticos
	QRY.cursorLocation 	= 3 					'cursor no cliente
	QRY.open cmd	

'VERIFICA SE É FINAL DE ARQUIVO
IF QRY.EOF THEN
	wstatus				    = "no"
ELSE
	wstatus 				= "yes"
	QRY.PageSize			= RegPorPag 		'Numero de registros por página
	registros				= QRY.RecordCount
	total					= QRY.PageSize
	paginas					= QRY.pageCount
	registro_atual			= QRY.absolutePosition
	QRY.AbsolutePage 		= PagAtual
	TotPag 					= QRY.PageCount
END IF%>

<iframe name="frameaux" id="frameaux" width="1" height="1" frameborder="0"></iframe>
<br><br>
<div id="mostraForm" style="width:98% !important; margin-left:15px;">    
     
	<form name="forme" id="forme" action="<%=wprograma%>.asp" method="post" class="form-inline" role="form">	
    <table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
    <tr>   
        <td>
	        <label class="control-label">NOME:</label><br>
           	<input type="text" name="irm_nome" id="irm_nome" value="<%=irm_nome%>" maxlength="70" class="form-control" placeholder="NOME.." title="Digite o nome!" style="width:300px">
        </td>
        <td width="10"></td>
        <td>
            <label class="control-label">MÊS REF.:</label><br>
            <input type="text" name="irm_dtNascimento" id="irm_dtNascimento" value="<%=irm_dtNascimento%>" maxlength="2" class="form-control" placeholder="DATA" title="Digite a data!" onKeyPress="return somenteNumeros(event)" onBlur="verifDataMesAno(this.value)" parametro="no" requerido="no" msg=""  style="width:150px">
        </td>
        <td width="10"></td>        
        <td>
        	<br>	
            <button type="submit" id="buscar" name="buscar" value="1" style="background:#003d7b;font-weight:bold;font-size:15px" class="btn btn-md btn-primary btn-block"> BUSCAR</button> 
		</td>
        <td width="10"></td>
        <td>
        	<br>	
            <button type="button" id="reativar" name="reativar" onClick="reativarAniver()" style="background:#003d7b;font-weight:bold;font-size:15px" class="btn btn-md btn-primary btn-block"> REATIVAR ANIVERSÁRIOS?</button> 
		</td>
	</tr>
    <tr>
    	<td colspan="10" style="font-size:9px; padding-top:3px; font-weight:bold"><font style="color:red">*</font>Campos livre para buscar aniversariantes de qualquer data!</td>
    </tr>
    </table>                                                           
	</form>
    
    <br>
    
<%IF wstatus = "yes" THEN%>    	    

    <div class="table-responsive-sm">
    
		<%IF wstatus = "yes" THEN%>                       
           <table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
           <tr>
           	    <td colspan="6" style="font-size:16px; text-align:center; font-weight:bold">ANIVERSÁRIOS DO MÊS</td>
           </tr>            
           </table>
            <br>
           <table class="table table-striped" style="font-size:12px" border="0">           
           <thead>
           		<tr>
                	<th scope="col"></th>  
                    <th scope="col" style="text-align:center">NOME</th>  
                    <th scope="col" style="text-align:center">IGREJA</th>                    
                    <th scope="col" style="text-align:center">DATA</th>
                    <th scope="col" style="text-align:center">IDADE EM <%=year(now)%></th>
                    <th scope="col" style="text-align:center">BAIXAR ANIVERSÁRIO?</th>                                                     
	            </tr>              
            </thead>
            
            <%'LOOP 
			DO WHILE count < RegPorPag AND NOT QRY.EOF
			count		= count + 1	

				'RECUPERA CAMPOS
				irm_cod					= QRY("irm_cod")
				irm_nome				= TRIM(UCASE(QRY("irm_nome")))
				irm_dtNascimento		= formataData(QRY("irm_dtNascimento"))
				igr_nFantasia			= TRIM(UCASE(QRY("igr_nFantasia")))
				irm_baixaNiver			= TRIM(UCASE(QRY("irm_baixaNiver")))%>
            	
                <tbody>
                    <tr>
                    	<th scope="col"></th> 
                        <td style="text-align:center"><%=irm_nome%></td>
                        <td style="text-align:center"><%=igr_nFantasia%></td>
                        <td style="text-align:center"><%=irm_dtNascimento%></td>
                        <td style="text-align:center"><%=year(now)-INT(RIGHT(irm_dtNascimento,4))%></td>
                        <%IF irm_baixaNiver <> "S" THEN%>
	                        <td align="center"><button class="btn btn-md btn-primary btn-block" onClick="baixaNiver('<%=irm_cod%>')" style="width:100px;text-align:center">SIM</button></td>
						<%ELSE%>
                        	<td align="center">BAIXADO</td>
                        <%END IF%>                            
                    </tr>
				</tbody>
            
            <%QRY.MOVENEXT
            LOOP
            QRY.CLOSE			:	SET QRY 		= nothing%>         
    	
        <%END IF%>
        
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
//fazer reativacao de todos os aniversario
function reativarAniver(){
	
	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#avisoConfirma').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#avisoConfirma').modal('show')
	$('#avisoTituloConfirma').text('Aviso do Sistema')
	$('#avisoDescricaoConfirma').text('Deseja realmente REATIVAR TODOS os aniversários? (Se o ano corrente mudou e os aniversários não foram reativados, clique em SIM)')
	$('#avisoImagemConfirma').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	$('#avisoConfirma').ready(function(e) {
         $('#nao').focus();	
    });				
    
	//foca o campo quando apertar ok
	$("#sim").click(function(){
		location='<%=wprograma%>.asp?irm_cod=<%=irm_cod%>&tipo=2'
	});	
	
}

// dar baixa no aniversario
function baixaNiver(irm_cod){

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#avisoConfirma').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#avisoConfirma').modal('show')
	$('#avisoTituloConfirma').text('Aviso do Sistema')
	$('#avisoDescricaoConfirma').text('Deseja realmente dar baixa nesse aniversário?')
	$('#avisoImagemConfirma').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	$('#avisoConfirma').ready(function(e) {
         $('#nao').focus();	
    });				
    
	//foca o campo quando apertar ok
	$("#sim").click(function(){
		location='<%=wprograma%>.asp?irm_cod='+irm_cod+'&tipo=1'
	});	
		
}

//coloca barra
function barraMesAno(valor){
	
	if (valor.length == 2){
		$("#irm_dtNascimento").val(valor+"/")
	}
}

//verif valor
function verifDataMesAno(valor){
	
	if (valor.substring(0,2) > 12 || valor.substring(0,2) < 1){
		$("#irm_dtNascimento").val("")
	}
}

function excluirUsuario(codigo){
	
	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#avisoConfirma').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#avisoConfirma').modal('show')
	$('#avisoTituloConfirma').text('Aviso do Sistema')
	$('#avisoDescricaoConfirma').text('Deseja realmente excluir esse usuário?')
	$('#avisoImagemConfirma').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	$('#avisoConfirma').ready(function(e) {
         $('#nao').focus();	
    });				
    
	//foca o campo quando apertar ok
	$("#sim").click(function(){
		frameaux.location = 'ajax/iframe_geral.asp?tipo=2&usu_cod=' + codigo
	});	
}
</script>

<script>
<%IF WERRO = 1 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Baixa realizada com sucesso!')
	$('#avisoImagem').addClass('glyphicon-ok text-success');
	
	// foco no botao
	$('#aviso').ready(function(e) {
         $('#ok').focus();	
    });				    		
	
<%END IF%>

<%IF WERRO = 2 THEN%>

	// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
	$('#aviso').modal({backdrop: 'static', keyboard: false}) 
	
	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text('Reativação de aniversários realizada com sucesso!')
	$('#avisoImagem').addClass('glyphicon-ok text-success');
	
	// foco no botao
	$('#aviso').ready(function(e) {
         $('#ok').focus();	
    });				    		
	
<%END IF%>
</script>

<%'FECHA TODAS AS QRY'S E CONEXÕES COM O BANCO DE DADOS
conexao.close			:	SET conexao 		= nothing%>