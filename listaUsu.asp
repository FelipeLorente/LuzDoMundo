<!--#include file="include/conexao.asp"-->
<!--#include file="include/topoPaginacaoUsuario.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<%'PARAMETROS OBRIGATORIOS
wprograma 		= "listaUsu"
usu_nome  		= request.Form("usu_nome")
usu_login		= request.Form("usu_login")
usu_email		= request.Form("usu_email")

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
'coloca o caminho da atualização
URL = wprograma&".asp"
		
'RECUPERA VALORES DO BANCO
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conexao
cmd.CommandText = "pr_le_usuariosBusca"
cmd.CommandType = 4
Set params = cmd.Parameters

params.Append cmd.CreateParameter("RETURN_VALUE", 3, 4, 0)	
params.Append cmd.CreateParameter("@usu_nome", 200, 1, 70, stringsqlserver(usu_nome))
params.Append cmd.CreateParameter("@usu_login", 200, 1, 50, stringsqlserver(usu_login))
params.Append cmd.CreateParameter("@usu_email", 200, 1, 100, stringsqlserver(usu_email))		

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
END IF

'FORMATA CNPJ
IF LEN(TRIM(igr_cnpj)) = 14 THEN
	igr_cnpj = FormataCNPJ(igr_cnpj)
ELSE
	igr_cnpj = ""
END IF%>

<iframe name="frameaux" id="frameaux" width="1" height="1" frameborder="0"></iframe>

<div id="mostraForm" style="width:98% !important; margin-left:15px;">    
     
	<form name="forme" id="forme" action="<%=wprograma%>.asp" method="post" class="form-inline" role="form">	
    <table cellpadding="2" cellspacing="0" border="0" align="center" style="font-size:12px">
    <tr>   
        <td>
	        <label class="control-label">NOME:</label><br>
           	<input type="text" name="usu_nome" id="usu_nome" value="<%=usu_nome%>" maxlength="70" class="form-control" placeholder="NOME 1.." title="Digite o nome!" style="width:300px">
        </td>
        <td width="10"></td>
        <td>
            <label class="control-label">LOGIN:</label><br>
            <input type="text" name="usu_login" id="usu_login" value="<%=usu_login%>" maxlength="50" class="form-control text-uppercase" placeholder="LOGIN 1.." title="Digite o login!" parametro="no" requerido="no" msg=""  style="width:150px">
        </td>
        <td width="10"></td> 
        <td>
            <label class="control-label">E-MAIL:</label><br>
            <input type="text" name="usu_email" id="usu_email" value="<%=usu_email%>" maxlength="20" class="form-control text-uppercase" placeholder="E-MAIL.." title="Digite o e-mail!" parametro="no" requerido="no" msg=""  style="width:300px">
        </td>
        <td width="10"></td>        
        <td>
        	<br>	
            <button type="submit" id="buscar" name="buscar" value="1" style="background:#003d7b;font-weight:bold;font-size:15px" class="btn btn-md btn-primary btn-block"> BUSCAR</button> </td>
	</tr>
    </table>                                                           
	</form>
    	    
    <div class="table-responsive-sm">
    
		<%IF wstatus = "yes" THEN%>                       
        
           <table class="table table-striped" style="font-size:12px">
           <thead>
           		<tr>
                    <th scope="col">NOME</th>                    
                    <th scope="col">LOGIN</th>
                    <th scope="col">E-MAIL</th>                    
                    <th scope="col">STATUS</th>
                    <th scope="col" title="Excluír este arquivos!">EXCLUÍR?</th>                                                      
	            </tr>              
            </thead>
            
            <%'LOOP 
			DO WHILE count < RegPorPag AND NOT QRY.EOF
			count		= count + 1	

				'RECUPERA CAMPOS
				usu_cod					= TRIM(UCASE(QRY("usu_cod"))) 
				usu_nome				= TRIM(UCASE(QRY("usu_nome")))
				usu_login				= TRIM(UCASE(QRY("usu_login")))
				usu_senha				= TRIM(UCASE(QRY("usu_senha")))
				usu_email				= TRIM(UCASE(QRY("usu_email")))
				usu_status				= TRIM(UCASE(QRY("usu_status")))
				usu_dtCadastro			= TRIM(UCASE(QRY("usu_dtCadastro"))) 
				usu_dtAlteracao			= TRIM(UCASE(QRY("usu_dtAlteracao"))) 				 
				
				'MONTA ENDEREÇO
				endereco = "NOME: " & usu_nome & ", LOGIN: " & usu_login & " E-MAIL: " & usu_email & " (" & retornastatus(usu_status) & ")"
				
				'MONTA LOCATION
				locationIgr = "parent.location='usuarios.asp?altera=sim&usu_cod="&usu_cod&"'"%>
            	
                <tbody>
                    <tr title="<%="Alterar? | " & UCASE(endereco)%>" style="cursor:pointer">
                        <td onClick="<%=locationIgr%>"><%=usu_nome%></td>
                        <td onClick="<%=locationIgr%>"><%=usu_login%></td>
                        <td onClick="<%=locationIgr%>" title="Enviar e-mail para <%=usu_nome%>?"><a href="mailto:<%=usu_email%>" style="color:black;"><%=usu_email%></a></td>
                        <td onClick="<%=locationIgr%>"><%=retornastatus(usu_status)%></td>
                        <td title="Deseja excluír este arquivo?" onClick="excluirUsuario('<%=usu_cod%>')" style="font-size:14px; font-weight:bold; color:red; padding-left:30px">X</td>
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

</div>

</body>
</html>

<!--#include file="include/bottom.asp"-->

<script>
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

<%'FECHA TODAS AS QRY'S E CONEXÕES COM O BANCO DE DADOS
conexao.close			:	SET conexao 		= nothing%>