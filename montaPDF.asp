<!--#include file="include/conexao.asp"-->
<!--#include file="include/topo.asp"-->
<!--#include file="include/expiraSessao.asp"-->

<%'PARAMETROS OBRIGATORIOS
wprograma 	= "montaPDF"
codigo		= Request.QueryString("abc")

'BUSCA IRMAO
SET QRY = conexao.execute("pr_le_irmaos 1, "&codigo&"")

'VERIF SE É FINAL DE ARQUIVO
IF NOT QRY.EOF THEN

	wstatus 					= "yes"
	
	'RECUPERA CAMPOS
	igr_cod 					= QRY("igr_cod") 
	igr_cnpj 					= QRY("igr_cnpj") 
	igr_rSocial 				= QRY("igr_rSocial") 
	igr_nFantasia 				= QRY("igr_nFantasia") 
	igr_cep 					= QRY("igr_cep") 
	igr_logradouro 				= QRY("igr_logradouro") 
	igr_numero 					= QRY("igr_numero") 
	igr_complemento 			= QRY("igr_complemento")  
	igr_bairro 					= QRY("igr_bairro")
	igr_cidade 					= QRY("igr_cidade") 
	igr_estado 					= QRY("igr_estado") 
	igr_internacional 			= QRY("igr_internacional") 
	igr_status 					= QRY("igr_status") 
	igr_responsavel 			= QRY("igr_responsavel") 
	igr_telResponsalvel 		= QRY("igr_telResponsalvel") 
	igr_dtCadastro 				= QRY("igr_dtCadastro") 
	igr_dtAlteracao				= QRY("igr_dtAlteracao")
	irm_nome 					= QRY("irm_nome") 
	irm_nPai 					= QRY("irm_nPai") 
	irm_nMae 					= QRY("irm_nMae") 
	irm_cep 					= QRY("irm_cep") 
	irm_logradouro 				= QRY("irm_logradouro")  
	irm_numero 					= QRY("irm_numero") 
	irm_complemento 			= QRY("irm_complemento") 
	irm_bairro 					= QRY("irm_bairro") 
	irm_cidade 					= QRY("irm_cidade") 
	irm_estado 					= QRY("irm_estado")
	irm_tel1 					= QRY("irm_tel1") 
	irm_tel2 					= QRY("irm_tel2") 
	irm_tpCel1 					= QRY("irm_tpCel1")
	irm_cel1 					= QRY("irm_cel1") 
	irm_tpCel2 					= QRY("irm_tpCel2") 
	irm_cel2 					= QRY("irm_cel2") 
	irm_nacionalidade 			= QRY("irm_nacionalidade") 
	irm_estCivil 				= QRY("irm_estCivil")
	irm_nConjegue 				= QRY("irm_nConjegue") 
	irm_qtdFilhos 				= QRY("irm_qtdFilhos") 
	irm_rg 						= QRY("irm_rg") 
	irm_cpf 					= QRY("irm_cpf") 
	irm_crente 					= QRY("irm_crente") 
	irm_btEspiritos 			= QRY("irm_btEspiritos") 
	irm_dtBtEspiritos 			= QRY("irm_dtBtEspiritos")
	irm_btAguas 				= QRY("irm_btAguas") 
	irm_dtBtAguas 				= QRY("irm_dtBtAguas") 
	irm_denominacaoBatismo 		= QRY("irm_denominacaoBatismo") 
	irm_CidEstPais 				= QRY("irm_CidEstPais") 
	irm_dtIngresso 				= QRY("irm_dtIngresso") 
	irm_funcIgreja 				= QRY("irm_funcIgreja") 
	irm_dtNascimento 			= QRY("irm_dtNascimento") 
	irm_obs  					= QRY("irm_obs")
	
	QRY.close			:	SET QRY 		= nothing
	conexao.close			:	SET conexao 		= nothing
END IF

IF wstatus = "yes" THEN
	
	theURL = ""
	'theURL = theURL & "<style>"
	'theURL = theURL & "		body{padding-bottom:0px;padding-top:10px}"
	'theURL = theURL & "</style>"
	
	theURL = theURL & "<table width='1000' border='0' align='center' cellpadding='0' cellspacing='0'>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td width='150'><img src='imagens/logoLuzDoMundo.jpg' width='150' height='150'></td>"
	theURL = theURL & "		<td width='10'></td>"
	theURL = theURL & "		<td width='790' valign='top'>"
	
	theURL = theURL & "			<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
	theURL = theURL & "			<tr height='30'>"
	theURL = theURL & "				<td colspan='99'></td>"
	theURL = theURL & " 		</tr>"
	theURL = theURL & "			<tr>"
	theURL = theURL & "				<td colspan='99'>CNPJ: "&FormataCNPJ(igr_cnpj)&" | <b>IGREJA: "&igr_rSocial&"</b></td>"
	theURL = theURL & "			</tr>"
	theURL = theURL & "			<tr height='10'>"
	theURL = theURL & "				<td colspan='99'></td>"
	theURL = theURL & "			</tr>"
	theURL = theURL & "			<tr>"
	theURL = theURL & "				<td colspan='99' style='font-size:12px'>ENDEREÇO: "&igr_logradouro&", "&igr_numero&" - "&igr_estado&"</td>"
	theURL = theURL & "			</tr>"
	theURL = theURL & "			<tr height='10'>"
	theURL = theURL & "				<td colspan='99'></td>"
	theURL = theURL & "			</tr>"
	theURL = theURL & "			<tr>"
	theURL = theURL & "				<td colspan='99' style='font-size:12px'>RESPONSÁVEL: "&igr_responsavel&" | "&igr_telResponsalvel&"</td>"
	theURL = theURL & "			</tr>"
	theURL = theURL & "			<tr height='10'>"
	theURL = theURL & "				<td colspan='99'></td>"
	theURL = theURL & "			</tr>"	
	theURL = theURL & "			<tr>"
	theURL = theURL & "				<td colspan='99' style='font-size:12px'>DATA DO CADASTRO: "&FORMATADATA(igr_dtCadastro)&"</td>"
	theURL = theURL & "			</tr>"
	theURL = theURL & "			</table>"
	theURL = theURL & "		</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "</table>"
	
	theURL = theURL & "<br>"
	
	theURL = theURL & "<table width='1000' border='0' align='center' cellpadding='0' cellspacing='0' style='font-size:13px; border:1px solid black;'>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='99' style='padding-left:10px'><b>NOME:</b> "&irm_nome&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99' style='padding-left:10px'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='99' style='padding-left:10px'><b>MÃE:</b> "&irm_nMae&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='99' style='padding-left:10px'><b>PAI:</b> "&irm_nPai&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td style='padding-left:10px'><b>TEL 1:</b> "&irm_tel1&"</td>"
	theURL = theURL & "		<td><b>TEL 2:</b> "&irm_tel2&"</td>"
	theURL = theURL & "		<td><b>CEL 1:</b> "&irm_tpCel1&" <b>|</b> "&irm_Cel1&"</td>"
	theURL = theURL & "		<td colspan='2'><b>CEL 2:</b> "&irm_tpCel2&" <b>|</b> "&irm_Cel2&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td style='padding-left:10px'><b>ENDEREÇO: </b> "&irm_logradouro&"</td>" 
	theURL = theURL & "		<td><b>NÚM: </b>"&irm_numero&"</td>"
	theURL = theURL & "		<td><b>COMP: </b>"&irm_complemento&"</td>"
	theURL = theURL & "		<td><b>BAIRRO: </b>"&irm_bairro&"</td>"
	theURL = theURL & "		<td><b>CIDADE: </b>"&irm_cidade&"</td>"
	theURL = theURL & "		<td><b>ESTADO: </b>"&irm_estado&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='2' style='padding-left:10px'><b>CRENTE: </b> "&retornastatus(irm_crente)&"</td>"
	theURL = theURL & "		<td colspan='10'><b>BATIZADO ESPÍRITO: </b>"&retornastatus(irm_btEspiritos)&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='2' style='padding-left:10px'><b>BATIZADO ÁGUAS: </b>"&irm_btAguas&" <b>DATA: </b>"&irm_dtBtAguas&"</td>"
	theURL = theURL & "		<td><b>CIDADE/ESTADO/PAÍS: </b>"&irm_CidEstPais&"</td>"
	theURL = theURL & "		<td colspan='99'><b>INGRESSO: </b>"&irm_dtIngresso&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='99' style='padding-left:10px'><b>DENOMINAÇÃO: </b> "&irm_denominacaoBatismo&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='99' style='padding-left:10px'><b>FUNÇÃO NA IGREJA: </b> "&irm_funcIgreja&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr height='10' style='border-top:1px dashed black'>"
	theURL = theURL & "		<td colspan='99'></td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "<tr>"
	theURL = theURL & "		<td colspan='99' style='padding-left:10px'><b>OBSERVAÇÕES: </b> "&irm_obs&"</td>"
	theURL = theURL & "</tr>"
	theURL = theURL & "</table>"

	response.Write theURL		
	response.End()
END IF%>