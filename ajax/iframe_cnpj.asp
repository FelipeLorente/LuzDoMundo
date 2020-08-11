<!--#include file="../include/conexao.asp"-->

<%'RECUPERA PARAMETROS
cnpj			= request.QueryString("cnpj")

'REMOVE BARRAS E PONTOS
cnpj = REPLACE(REPLACE(REPLACE(cnpj,".",""),"-",""),"/","")%>

<%'VERIFICA SE O CNPJ E VALIDO
verifCNPJFiliais = CalculaCNPJ(cnpj)

'VERIF FILIAIS
IF verifCNPJFiliais = 1 THEN%>
	
    <script>
		// mostra aviso de erro				
		parent.$('#avisoImagem').removeClass('glyphicon glyphicon-ok')	
		parent.$('#avisoImagem').addClass('glyphicon glyphicon-remove')	
		parent.$('#avisoImagem').css('color' , 'red')	
		parent.$('#avisoClass').removeClass('modal-lg')
		parent.$('#avisoClass').addClass('modal-sm')
		
    	parent.$('#aviso').modal('show')
		parent.$('#avisoTitulo').text('Aviso do Sistema')
		parent.$('#avisoDescricao').text('CNPJ inválido!')
		parent.$('#camposenha').addClass('hidden')
		parent.$('#senha').val('')	
		
		// foco no botao
		parent.$('#aviso').on('shown.bs.modal', function(){
		  parent.$('#ok').focus();
		});
				
		// limpa campo		
		parent.$("#igr_cnpj").val('')			
			
    </script>

<%END IF%>

<!--#include file="../include/fechaConexao.asp"-->
