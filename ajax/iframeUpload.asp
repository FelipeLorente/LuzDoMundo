<!--#include file="../include/conexao.asp"-->
<!--#include file="../include/expiraSessao.asp"-->

<%'PARAMETROS OBRIGATORIOS
wprograma  	  = "iframeUpload"
enviar		  = Request.QueryString("enviar")
campo		  = Request.QueryString("campo")

'VERIFICA SE TEM TIPO
IF NOT ISEMPTY (enviar) THEN

	'VERIFICA A SESSION
	IF LEN(SESSION("age_codigo")) = 0 THEN
		REDIRECTPAGE(2)	
		response.End() 
	END IF	
	
	'VERIFICA SE TEM CAMPO
	IF LEN(TRIM(campo)) = 0 THEN
		
		'MOSTRA GIF DE LOADING
		response.Write "<script>parent.$('#loading7').removeClass('hidden')</script>"
	
		'PESQUISA NO BD QUAL O ULTIMO POSIÇÃO E SOMA 1
		'LISTA ANEXOS
		SET QRY = conexao.execute("pr_si_le_anexosagencias 3 , "&SESSION("age_codigo")&" , NULL , NULL ")
		
		'VERIFICA SE É FINAL DE ARQUIVO
		IF NOT QRY.EOF THEN
			'RECUPERA CAMPOS
			campo = QRY("ane_posicao") 
			
			QRY.close			:	SET QRY 		= nothing
			
			'VERIFICA SE PASSOU DAS POSIÇÕES FIXAS PARA SOMAR 1
			IF campo > 6 THEN
				campo = campo + 1
			ELSE
				campo = 7
			END IF
		ELSE
			campo = 7
		END IF
	END IF				

	' Coloque um número grande para o tempo de finalização do script, pois o upload pode demorar alguns minutos. Se o servidor estiver com o tempo baixo, pode haver erro no upload 
	Server.scripttimeout = 999999
	
	' Caso houver algum erro o código vai prosseguir até o final. Isso evita que seja mostrada aquela página de erro padrão do Internet Explorer 
	On Error Resume Next 
	
	' Aqui criamos uma instância do objeto do ASP Smart Upload 
	Set Upload = Server.CreateObject("aspSmartUpload.SmartUpload") 		
	
	' Aqui criamos uma lista dos formatos de arquivos que poderão ser enviados 
	Upload.AllowedFilesList = "pdf" 
	
	' Aqui configuramos o tamanho máximo de cada arquivo enviado em bytes 
	Upload.MaxFileSize = 16000000  'de 970 a 1000 kb
	
	' Aqui configuramos o tamanho total para os arquivos enviados. Todos os arquivos juntos não podem passar deste tamanho 
	Upload.TotalMaxFileSize = 40000000 
	
	' pasta onde sera armazenado
	pasta		 = server.MapPath("..\arquivos")&"\"
	
	' Tipo de arquivo que esta sendo enviado
	tamanho = round(request.TotalBytes/1024)
	
	'retira tracos e pontos
	cnpj = SESSION("age_cnpj")
			
	' Aqui é efetuado o envio dos arquivos 
	Upload.Upload 
	
	'declaração de variaveis
	arquivo 		= Upload.Form("ane_url").values
	'CRIA OBJETO PARA MANIPULAR OS ARQUIVOS
	Set oFS = Server.CreateObject("Scripting.FileSystemObject")
				
	nome_arquivo 	= campo&"_"&cnpj&"_COD"&SESSION("age_codigo")&".pdf"	

	 'Se houver algum erro será exibida essa mensagem e a descrição do erro 
	 If Err Then 
		IF REPLACE(TRIM(ERR.NUMBER),"-","") = "2147220399" THEN			
			 response.write "<script>parent.chamaAlerta('O arquivo excedeu o tamanho permitido de 15mb!', '', 'etapa5.asp'); </script>"
		ELSEIF REPLACE(TRIM(ERR.NUMBER),"-","") = "2147220494"  THEN
			response.write "<script>parent.chamaAlerta('Somente o formato PDF é aceito!', '', 'etapa5.asp'); </script>"
		END IF
	 End if
	 									
	' Selecionamos cada arquivo que foi submetido do formulário 
	For each File in Upload.Files 	
		
		' Aqui checamos se o tamanho dele é maior que 0 byte. Isso é necessário pois se a pessoa submeter o formulário com o endereço do arquivo errado, será criado um 
		If File.Size > 0 Then
		
			If LEN(File.FileName) > 0 Then									 							
				
				
				'VERIFICA SE TEM ANEXO CADASTRADO
				SET QRY_ANEXO = conexao.execute("pr_si_le_anexosagencias 4,"&session("age_codigo")&", NULL, '"&nome_arquivo&"' ")								
				
				'VERIFICA FINAL DE ARQ
				IF QRY_ANEXO.EOF THEN								
					'CADASTRA ARQUIVO COM COD DA TALBELA AGENCIAS  
					conexao.execute("pr_si_ma_anexosagencias 1,"&session("age_codigo")&", '"&nome_arquivo&"', "&campo&", NULL ")
				END IF								
																							
				'salva de acordo com o nome desejado
				File.SaveAs(pasta&nome_arquivo)
				
				'ATRIBUI VALOR 0 PARA O ERRO
				ERRO = 0										
				
			end if
		'Caso for um arquivo inválido, ou seja, o tamanho dele for igual a zero então aparecerá a mensagem e em seguida terminamos a condição. 
		 Else 		 
			response.write "<script>parent.chamaAlerta('Arquivo vazio não é permitido!', '', 'etapa5.asp');</script>"   
		 End if 
	 		 
	'Caso mais de um arquivo tenha sido enviado, enviamos o sistema para o próximo. 
	 Next 	
	 
	 IF LEN(erro) > 0 THEN%>
		<script>
        
		// muda caracteristica do modal deixa o modal estatico, a tela pai fica inabilitada
        parent.$('#aviso').modal({backdrop: 'static', keyboard: false}) 
        
        // mostra aviso de erro
        parent.$('#aviso').modal('show')
        parent.$('#avisoTitulo').text('Aviso do Sistema')
        parent.$('#avisoDescricao').text('Arquivo salvo com sucesso!')
        parent.$('#avisoImagem').addClass('glyphicon-ok text-danger');
        
        // foco no botao
        parent.$('#aviso').ready(function(e) {
             parent.$('#ok').focus();	
        });				
        
        //redireciona para prox etapa
        parent.$("#ok").click(function(){
            parent.$(parent.location).attr('href', 'etapa5.asp')
        });	
        
        </script>
    <%END IF
	 			
END IF%>

<!--#include file="../include/fechaConexao.asp"-->