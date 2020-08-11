<%@ LANGUAGE = VBScript.Encode %>
<%'expira tudo o que estava armazenado
response.Buffer 		= true
response.addHeader 		"pragma", "no-cache"
response.CacheControl 	= "Private"
response.Expires 		= 0
session.TimeOut			= 1440
session.LCID 			= 1033
server.scripttimeout    = 1440
'on error resume next

'INDENTIFICA OS MESSES
IF month(date()) = 1 THEN
	mes = "Janeiro"
ELSEIF month(date()) = 2 THEN
	mes = "Fevereiro"
ELSEIF month(date()) = 3 THEN
	mes = "Março"
ELSEIF month(date()) = 4 THEN
	mes = "Abril"
ELSEIF month(date()) = 5 THEN
	mes = "Maio"
ELSEIF month(date()) = 6 THEN
	mes = "Junho"
ELSEIF month(date()) = 7 THEN
	mes = "Julho"
ELSEIF month(date()) = 8 THEN
	mes = "Agosto"
ELSEIF month(date()) = 9 THEN
	mes = "Setembro"
ELSEIF month(date()) = 10 THEN
	mes = "Outubro"
ELSEIF month(date()) = 11 THEN
	mes = "Novembro"
ELSEIF month(date()) = 12 THEN
	mes = "Desembro"
END IF

'INDENTIFICA OS MESSES
function verificaMes(d)

	IF d = "1" THEN
		verificaMes = "Janeiro"
	ELSEIF d = "2" THEN
		verificaMes = "Fevereiro"
	ELSEIF d = "3" THEN
		verificaMes = "Março"
	ELSEIF d = "4" THEN
		verificaMes = "Abril"
	ELSEIF d = "5" THEN
		verificaMes = "Maio"
	ELSEIF d = "6" THEN
		verificaMes = "Junho"
	ELSEIF d = "7" THEN
		verificaMes = "Julho"
	ELSEIF d = "8" THEN
		verificaMes = "Agosto"
	ELSEIF d = "9" THEN
		verificaMes = "Setembro"
	ELSEIF d = "10" THEN
		verificaMes = "Outubro"
	ELSEIF d = "11" THEN
		verificaMes = "Novembro"
	ELSEIF d = "12" THEN
		verificaMes = "Dezembro"
	END IF
end function

'coloca número para o padrão sql
function stringsqlserver(s)
	IF not isnull(s) and len(s) > 0 THEN
		s			= replace(s, "'", "")
		stringsqlserver	= s
	ELSE
		stringsqlserver = NULL
	END IF
end function

'sql injection
function sqlInjection(v)
	var	= trim(v)
	var = replace(var,"insert into","")
	var = replace(var,"select","")
	var = replace(var,"*","")
	var = replace(var,"'","")
	var = replace(var,"delete","")
	var = replace(var,"update","")
	var = replace(var,"insert","")
	var = replace(var,"set","")
	var = replace(var,"(","")
	var = replace(var,")","")		
	
	sqlInjection = var
end function

'MOSTRA DA COMPLETA agosto 31, 1985
MinhaData = date()

'DEIXA PRIMEIRA LETRA MAIUSCULA
FUNCTION letraMiuscula(valor)
	PLetra = UCASE(left(valor,1))
	resto  = len(valor)-len(PLetra)
	restante = LCASE(right(valor,resto))
	
	'COLOCA TODAS AS PRIMEIRAS LETRAS MAIUSCULAS
	valor = PLetra&restante
	
	arrayEsp = Split(valor, " ", -1, 1)
	
	for i = 0 to UBOUND(arrayEsp)
		letraM   = UCASE(left(arrayEsp(i),1))
		restoM   = len(arrayEsp(i))-len(letraM)
		restanteM = LCASE(right(arrayEsp(i),restoM))
		
		total = letraM&restanteM
		
		if InstrRev(lcase(total), "ii", -1, 1)>0 or InstrRev(lcase(total), "/", -1, 1)>0 then 
			total = UCASE(total) 
		else
			total = total
		end if
		
		if i = 0 then
			valorM = total
		else
			valorM = valorM&" "&total
		end if

	next
	
	if len(valorM)>0 then
		letraMiuscula = valorM
	else
		letraMiuscula = PLetra&restante
	end if
END FUNCTION

'MONTA DATA DA SEMANA
FUNCTION MontaDataExtenso()

	dia = day(now)
	mes = month(now)
	ano = year(now)
	
	IF LEN(dia) = 1 THEN
		dia = "0"&dia
	END IF
	
	MontaDataExtenso = dia & " de " & verificaMes(mes) & " de " & ano
END FUNCTION

'INDENTIFICA OS DIAS DA SEMANDA SEG,TER,QUAR,QUIN,SEX,SÁB,DOMI
IF trim(cint(WeekDay(MinhaData))) = 1 THEN
	semana = "Domingo"
ELSEIF trim(cint(WeekDay(MinhaData))) = 2 THEN
	semana = "Segunda"
ELSEIF trim(cint(WeekDay(MinhaData))) = 3 THEN
	semana = "Terça"
ELSEIF trim(cint(WeekDay(MinhaData))) = 4 THEN
	semana = "Quarta"
ELSEIF trim(cint(WeekDay(MinhaData))) = 5 THEN
	semana = "Quinta"
ELSEIF trim(cint(WeekDay(MinhaData))) = 6 THEN
	semana = "Sexta"
ELSEIF trim(cint(WeekDay(MinhaData))) = 7 THEN
	semana = "Sábado"
END IF

'INDENTIFICA PERIODO DE UM DIA
IF HOUR(NOW) >= 7 AND HOUR(NOW) < 12 THEN
	periodo = "Bom dia"
ELSEIF HOUR(NOW) >= 12 AND HOUR(NOW) < 18 THEN
	periodo = "Boa tarde"
ELSEIF HOUR(NOW) >= 18 AND HOUR(NOW) < 24 THEN
	periodo = "Boa noite"
ELSEIF HOUR(NOW) >= 24 AND HOUR(NOW) < 7 THEN
	periodo = "Boa madrugada"
END IF

'coloca valor para o padrão sql
Function FormataMoeda(valor)
  sp = Mid(FormatNumber(1000, 0, -1, 0, -1),2,1)
  sv = Mid(FormatNumber(0.1, 1, -1, 0, -1),2,1)
  If isNumeric(valor) Then
    v = FormatNumber (valor, 2, -1, 0, -1)
  Else
    v = FormatNumber (0, 2, -1, 0, -1)
  End If
  v = Replace(v, sp, "p")
  v = Replace(v, sv, "v")
  v = Replace(v, "p", ".")
  v = Replace(v, "v", ",")
  FormataMoeda = v
End Function

'retira todos os pontos, traco, etc
Function retiraCaracteres(valor)
   valor = replace(valor,"-","")
   valor = replace(valor,".","")
   valor = replace(valor,",","")
   valor = replace(valor,"/","")
   valor = replace(valor,"\","")
   valor = replace(valor,"@","")
   
   retiraCaracteres = valor
End function

'retira todos os acentos para pesquisar no banco de dados
Function retiraAcentoBD(valor)
    valor = Replace(valor, "á", "_")
	valor = Replace(valor, "à", "_")
	valor = Replace(valor, "â", "")
	valor = Replace(valor, "ã", "_")	
    valor = Replace(valor, "é", "_")
	valor = Replace(valor, "è", "_")
	valor = Replace(valor, "ê", "")	
    valor = Replace(valor, "í", "_")
	valor = Replace(valor, "ì", "_")
	valor = Replace(valor, "î", "")	
    valor = Replace(valor, "ó", "_")
	valor = Replace(valor, "ò", "_")
	valor = Replace(valor, "ô", "")
	valor = Replace(valor, "õ", "_")	
    valor = Replace(valor, "ú", "_")
	valor = Replace(valor, "ù", "_")
	valor = Replace(valor, "û", "")
   
   retiraAcentoBD = valor
End function


'formata cnpj
Function FormataCNPJ(Termo)
   Termo = Termo
   IF Len(Termo) >=14 Then
       Resultado = Resultado & Mid(Termo,1,2) & "." & Mid(Termo,3,3) & "." & Mid(Termo,6,3) & "/" & Mid(Termo,9,4) & "-" & Mid(Termo,13,2)
   Else
       Resultado = "CNPJ deve apresentar 14 dígitos!"
   End IF
   FormataCNPJ = Resultado
End function

'formata CEP
Function FormataCep(valor)
	letrasIni = mid(valor,1,5)
	letrasFim = mid(valor,6,8)
	
	FormataCep = letrasIni&"-"&letrasFim
End Function

'formata TEL
Function FormataTel(valor)
	IF valor>0 THEN
		letrasIni = mid(valor,1,4)
		letrasFim = mid(valor,5,8)
		
		FormataTel = letrasIni&"-"&letrasFim
	ELSE
		FormataTel = 0
	END IF
End Function

'formata SITE
Function FormataSite(valor)
	letras = InstrRev(lcase(valor), "http://", -1, 1)

	if letras = 0 then
		FormataSite = "http://"&valor
	else
		FormataSite = valor
	end if
	
End Function

'coloca valor para o padrão sql
Function FormataValorAmericano(valor)
  If isNumeric(valor) Then
	v = FormatCurrency (valor)
  Else
	v = 0
  End If
  v = Replace(v, "$", "R$ ")
  v = Replace(v, ",", "*")
  v = Replace(v, ".", ",")
  v = Replace(v, "*", ".")
  FormataValorAmericano = v
End Function

Function formataHora(DataHora)
  formataHora = Right("0" & DatePart("H",DataHora),2) & ":" & Right("0" & DatePart("n",DataHora),2) & ":" & Right("0" & DatePart("s",DataHora),2)
End Function

'gera numero aleatorio para resolver alguns problemas de cache
randomize
aleatorio = int(rnd * 999999999) + 0

'CONEXAO COM BD
Set conexao = Server.CreateObject("ADODB.Connection")
conexao.Open "Provider=SQLOLEDB.1;SERVER=LUZDOMUNDO\SQLEXPRESS;DATABASE=LuzDoMundo;UID=sa;PWD=sa"
	
FUNCTION retornastatus(s)
	IF 		s = "A" THEN 
		retornastatus = "ATIVO" 
	ELSEIF 	s = "D" THEN 
		retornastatus = "DESATIVADO"
	ELSEIF 	s = "TO" THEN 
		retornastatus = "Topo"
	ELSEIF 	s = "RO" THEN 
		retornastatus = "Rodapé"
	ELSEIF 	s = "S" THEN 
		retornastatus = "SIM"
	ELSEIF 	s = "N" THEN 
		retornastatus = "NAO"			
	END IF
END FUNCTION

FUNCTION retornavl(s)
	IF 		s = "S" THEN 
		retornavl = "Sim" 
	ELSEIF 	s = "N" THEN 
		retornavl = "Não"
	ELSEIF 	s = "H" THEN 
		retornavl = "Nihil"
	ELSEIF 	s = "T" THEN 
		retornavl = "Talvez"
	ELSEIF 	s = "P" THEN 
		retornavl = "Pendente"
	ELSEIF 	s = "F" THEN 
		retornavl = "Finalizado"
	END IF
END FUNCTION

'PREENCHE ZEROS À ESQUERDA
function zeros(str, num)
	WHILE LEN(str) < num
		str = "0"&str
	WEND
	
	zeros = str
end function


'retorna mensagens de acordo com o banco de dados
function retornamsg(mc)

	if isnumeric(mc) then
		if int(mc) < 0 then
			mc = mc * -1
		end if
	end if

	retornamsg = "Mensagem  no."&mc&"  não foi encontrada."
end function

'ordenar
function ordenar(cp)
	ordenar = "style=""cursor:hand"" onclick=""ordenar('"&cp&"')"" title=""Clique para ordenar esta lista nesta ordem."""
end function

'preenche totais de registro por página
function totalreg(t)
	response.write "<script type='text/javascript'>totalreg('"&t&"');</script>"
end function

'determina página atual da paginação
sub paginacao()

	if pag_acao = "proxima" then
		if int(pag_atual + 1) < int(pag_count) then
			pag_atual = int(pag_atual + 1)
		else
			pag_atual = pag_count
		end if
	elseif pag_acao = "anterior" then
		if int(pag_count) > int(pag_atual - 1) and int(pag_atual - 1) > 0 then
			pag_atual = int(pag_atual - 1)
		else
			pag_atual = 1
		end if
	elseif pag_acao = "pagina" then
		if isnumeric(pag_numero) then
			pag_atual = pag_numero
		else
			pag_atual = 1
		end if
	end if

end sub


'coloca data para o padrão sql
function datasql(s)
	IF isdate(s) THEN
		wdia 	= mid(s, 1, 2)
		wmes 	= mid(s, 4, 2)
		wano 	= mid(s, 7, 4)
		datasql = wano&"-"&wmes&"-"&wdia
	ELSE
		datasql = null
	END IF
end function


'coloca valor para o padrão sql
function valorsql(s)
	IF isnumeric(s) AND len(s) > 0 THEN
		valorsql = replace(s, ",", "")
	ELSE
		valorsql = null
	END IF
end function

'coloca número para o padrão sql
function numerosql(s)
	IF isnumeric(s) AND len(s) > 0 THEN
		numerosql = s
	ELSE
		numerosql = "0"
	END IF
end function

'coloca número para o padrão sql
function stringsql(s)

	IF not isnull(s) and len(s) > 0 THEN
		stringsql	= "'"&s&"'"
	ELSE
		stringsql = "''"
	END IF
end function

'deixa a primeira letra da String maiúscula
function maiuscula(s)
	c = len(s)
	if c > 0 then
		l = left(s,1)
		r = right(s, c-1)
		maiuscula = ucase(l)&lcase(r)	
	else
		maiuscula = s
	end if
end function

'formata a data para dd/mm/aaaa QUANDO LE
function formatadata(data)

	if not isnull(data)THEN
		if len(data) > 0 then
			if len(trim(day(data))) = 1 then
				fdia = "0"&day(data)
			else
				fdia = day(data)
			end if
			if len(trim(month(data))) = 1 then
				fmes = "0"&month(data)
			else
				fmes = month(data)
			end if
			
			formatadata = fdia&"/"&fmes&"/"&year(data)
			
			if formatadata = "//" then
				formatadata = ""
			end if
		else
			formatadata = ""
		end if
	else
		formatadata = ""
		
	end if			

end function

'formata a data para dd/mm/aaaa QUANDO LE
Function FormataDataAmericano(Data)
  v = Right("0" & Mid(Data,4, 2),2) & "/" & Right("0" & Left(Data, 2),2) & "/" & DatePart("yyyy", Data)
  v = Cdate(v)
  FormataDataAmericano = v
End Function

'formata a data para dd/mm/aaaa QUANDO LE
Function FormataDataBR(Data)
  v = Right(left(Data,10),2)&"/"&Mid(left(Data,10), 6, 2)&"/"&Left(left(Data,10),4)

  FormataDataBR = v
End Function

function montaEscala(datade,dataate)
	
	'VERIFICA SE É HORAS DE UMA DATA
	IF DateDiff("d", FormataDataAmericano(datade), FormataDataAmericano(dataate)) = 0 THEN
		
		'COLOCA O TIPO DE DATA
		tipoDT = "H"
		
		'CONTA QUANTOS DIAS TEM
		FOR x = 0 TO 23
								
			'verifica se o acumulador esta vazio
			IF LEN(xPERIODO)=0 THEN
				xPERIODO = x
			ELSE
				xPERIODO = xPERIODO&"*"&x
			END IF

		NEXT
	
	'VERIFICA SE É DIAS DE UMA DATA
	ELSEIF DateDiff("d", FormataDataAmericano(datade), FormataDataAmericano(dataate)) > 0 AND DateDiff("d", FormataDataAmericano(datade), FormataDataAmericano(dataate)) <= 30 THEN
		
		'COLOCA O TIPO DE DATA
		tipoDT = "D"
		
		diade  	= cint(day(FormataDataAmericano(datade)))*1
		diaate	= cint(day(FormataDataAmericano(dataate)))*1
					
		'CONTA QUANTOS DIAS TEM
		FOR x = 0 TO DateDiff("d", FormataDataAmericano(datade), FormataDataAmericano(dataate))
								
			'verifica se o acumulador esta vazio
			IF LEN(xPERIODO)=0 THEN
				xPERIODO = day(DateAdd("d",x, FormataDataAmericano(datade)))
			ELSE
				xPERIODO = xPERIODO&"*"&day(DateAdd("d",x, FormataDataAmericano(datade)))
			END IF

		NEXT
		
	'VERIFICA SE É MES DE UMA DATA
	ELSEIF DateDiff("d", FormataDataAmericano(datade), FormataDataAmericano(dataate)) > 30 AND DateDiff("d", FormataDataAmericano(datade), FormataDataAmericano(dataate)) <= 364 THEN
		
		'COLOCA O TIPO DE DATA
		tipoDT = "M"
		
		diade  	= cint(month(FormataDataAmericano(datade)))*1
		diaate	= cint(month(FormataDataAmericano(dataate)))*1
								
		'CONTA QUANTOS DIAS TEM
		FOR x = diade TO diaate
										
			'verifica se o acumulador esta vazio
			IF LEN(xPERIODO)=0 THEN
				xPERIODO = x
			ELSE
				xPERIODO = xPERIODO&"*"&x
			END IF

		NEXT
	
	'VERIFICA SE É ANO DE UMA DATA
	ELSEIF DateDiff("d", FormataDataAmericano(datade), FormataDataAmericano(dataate)) > 364 THEN
		
		'COLOCA O TIPO DE DATA
		tipoDT = "A"
		
		diade  	= cint(year(FormataDataAmericano(datade)))*1
		diaate	= cint(year(FormataDataAmericano(dataate)))*1
								
		'CONTA QUANTOS DIAS TEM
		FOR x = diade TO diaate
				
			'verifica se o acumulador esta vazio
			IF LEN(xPERIODO)=0 THEN
				xPERIODO = x
			ELSE
				xPERIODO = xPERIODO&"*"&x
			END IF

		NEXT	
	END IF
	
	montaEscala = tipoDT&";"&xPERIODO
	
end function

'TV-1,1687*2,61;1748||RV-1,0*2,0;0 (tipo-dia/mes/ano,total;dia/mes/ano,total;total geral)
function geraGrafico(caminho,tamanho,tipo,tituloGrafico,legenda,itens)			

'VERIFICA SE O CAMINHO NAO FOI DIGITADO
IF LEN(caminho)=0 THEN
	caminho = Server.MapPath("graficos")&"\"
END IF

'VERIFICA SE O TAMANHO DA IMAGEM FOI PREENCHIDO
IF LEN(tamanho)=0 THEN
	largura = 600
	altura  = 400 
ELSE
	tamanho = UCASE(tamanho)
	largura = trim(left(tamanho,InstrRev(tamanho, "X", -1, 1)-1))
	altura  = trim(right(tamanho,len(tamanho)-InstrRev(tamanho, "X", -1, 1)))

END IF

'VERIFICA SE O TITULO FOI DIGITADO
IF LEN(tituloGrafico)=0 THEN
	tituloGrafico = "EXEMPLO"
END IF

'VERIFICA SE O TITULO FOI DIGITADO
IF LEN(legenda)=0 THEN
	legenda = "LEGENDA"
END IF

'DESMENBRA OS ITESNS E VALORES
IF LEN(itens)=0 THEN
	response.write "Não há itens / valores para fazer a montagem do grafico!"
	response.End()
END IF

	'EXCLUI OS GRAFICOS DOS DIAS ANTERIORES
	'Declara as variáveis a serem usadas
	Dim Local, Folder, File, ObjFS, objRootFolder 
	
	'Especifica o endereço do conteúdo a ser exibido
	Local = request.ServerVariables("APPL_PHYSICAL_PATH")
	
	
	'Cria o objeto FileSystemObject
	Set ObjFS = Server.CreateObject("Scripting.FileSystemObject")
	Set objFolder = ObjFS.GetFolder(Server.MapPath("graficos")&"\")	
	
	'Exibe arquivos encontrados
	For Each File in objFolder.files
	
		'VERIFICA OS FORMATOS PARA TIRAR O .DB E O ARQUIVO PADRÃO
		IF RIGHT(UCASE(File.Name), 2) <> "DB" THEN
		
			'VERIFICA SE A DATA DE CRIAÇÃO É MENOR QUE HOJE
			IF  DateDiff("d", Now, File.DateLastModified) <0 THEN

				ObjFS.DeleteFile Server.MapPath("graficos")&"\"&File.Name
			
			END IF
			
		END IF
		
	Next		
	
	'Elimina variáveis da memória 
	Set Local = Nothing
	Set File = Nothing
	Set objFolder = Nothing
	Set Folder = Nothing
	
	With Server.CreateObject("OWC10.ChartSpace")
		set c = .constants
		With .charts.add
		
			'coloca borda e cor na borda
			'.Border.Color = "navy"
    		'.Border.Weight = c.owcLineWeightMedium
			
			.HasTitle=True
			if .HasTitle then
				.Title.Caption = tituloGrafico
				.Title.Font.Size = 14
				.Title.Font.Name = "Tahoma"
				.Title.Font.Bold = True
				.Title.Font.Italic = False
				.Title.Font.Underline = False
			end if
			
			.HasLegend = false
			if .HasLegend then
				.Legend.Font.Size = 12
				.Legend.Font.Name = "Tahoma"
				.Legend.Font.Bold = False
				.Legend.Font.Italic = False
				.Legend.Font.Underline = False
				.Legend.Position = c.chLegendPositionright
			end if
	
			Select Case tipo
			Case 1
				.Type = c.chChartTypeArea
			Case 2
				.Type = c.chChartTypeBarClustered
			Case 3
				.Type = c.chChartTypeColumnClustered
			Case 4
				.Type = c.chChartTypeLine
			Case 5
				.Type = c.chChartTypeLineMarkers
			Case 6
				.Type = c.chChartTypeColumn3d
			Case 7
				.Type = c.chChartTypePie
			Case 8
				.Type = c.chChartTypePie3d
			Case 9
				.Type = c.chChartTypeColumnClustered3D
			Case Default
				.Type = c.chChartTypePie3d
				'.Type = c.chChartTypeColumn3d
				'.Type = c.chChartTypeArea3D
				'.Type = c.chChartTypeBar3D
				'.Type = c.chChartTypeColumn3D
				'.Type = c.chChartTypeColumnStacked3D
				'.Type = c.chChartTypeCombo
				'.Type = c.chChartTypeCombo3D
			End Select
			
			'VERIFICA SE É APENAS UM ITEM
			IF InstrRev(itens, "||", -1, 1) = 0 THEN
			
				nome 	= UCASE(REPLACE(LEFT(itens,INSTRREV(itens, "-", -1, 1)-1),"'",""))
				resto 	= RIGHT(itens,LEN(itens)-InstrRev(itens, "-", -1, 1))
				
				'verifica se pegou o total 
				if InstrRev(resto, ";", -1, 1) <> 0 then
					totalGeral = Cdbl(trim(right(resto,len(resto)-InstrRev(resto, ";", -1, 1))))
					resto	   = left(resto,InstrRev(resto, ";", -1, 1)-1)
				end if
								
				ARRAYQTD = Split(resto, "*", -1, 1)
				
				acuValor   	= ""
				acuPeriodo 	= ""
				acuValorTot = ""
				
				for x = 0 to UBOUND(ARRAYQTD)
					qtdPeriodo  = Cdbl(trim(right(ARRAYQTD(x),len(ARRAYQTD(x))-InstrRev(ARRAYQTD(x), ",", -1, 1))))
					periodo 	= trim(left(ARRAYQTD(x),InstrRev(ARRAYQTD(x), ",", -1, 1)-1))
					
					'verifica se o valor nao é zero
					if qtdPeriodo >0 then
						
						conta = ((qtdPeriodo*100)/totalGeral)
						
						'ARRUMA O VALOR DAS PORCENTAGENS
						IF InstrRev(conta, "-", -1, 1) >0 THEN
						
							IF Right(conta, 1) = 2 THEN
								valor = "0.0"&REPLACE(left((conta), InstrRev(conta, ".", -1, 1)-1), ".", "")
								
							ELSEIF Right(conta, 1) = 1 THEN
								valor = "0."&REPLACE(left((conta), InstrRev(conta, ".", -1, 1)+1), ".", "")
							ELSE
								valor = left((conta), InstrRev(conta, ".", -1, 1)+2) 
							END IF
							
						ELSE
							valor = left((conta), InstrRev(conta, ".", -1, 1)+2) 
						END IF						
						
						if len(acuValor) = 0 then
							acuValor 	= valor&""
							acuValorTot = valor&"% ("&qtdPeriodo&") "
						else
							acuValor 	= acuValor&","&valor&""
							acuValorTot = acuValorTot&" / "&valor&"% ("&qtdPeriodo&")"
						end if
					else
						if len(acuValor) = 0 then
							acuValor 	= "0"
							acuValorTot = "0% ("&qtdPeriodo&") "
						else
							acuValor 	= acuValor&",0"
							acuValorTot = acuValorTot&" / 0% ("&qtdPeriodo&")"
						end if
					end if
					
					if len(acuPeriodo) = 0 then
						acuPeriodo = periodo
					else
						acuPeriodo = acuPeriodo&","&periodo
					end if
				next
						
				'monta grafico automaticamente
				With .seriescollection.add
					.DataLabelsCollection.Add
					.caption = nome &" - " &acuValorTot &" Total = "&totalGeral
					
					'so mostra o periodo uma vez
					'if i=0 then
						.setdata c.chdimcategories,c.chdataliteral, replace(acuPeriodo, "01", "1")
					'end if					
													
					.setdata c.chdimvalues,c.chdataliteral, acuValor
					'.Interior.Color = "green"
					
				End With
				
				'.Axes(c.chAxisPositionleft).NumberFormat = "#,##0"
				'.Axes(c.chAxisPositionleft).Scaling.Minimum = 0 
				'.Axes(c.chAxisPositionleft).Scaling.Maximum = 100
								
			ELSE
			
				ARRAYITENS = Split(itens, "||", -1, 1)
				
				'ADICIONA ITESN E VALORES
				for i = 0 to UBOUND(ARRAYITENS)
					nome 	= UCASE(REPLACE(LEFT(ARRAYITENS(i),INSTRREV(ARRAYITENS(i), "-", -1, 1)-1),"'",""))
					resto 	= RIGHT(ARRAYITENS(i),LEN(ARRAYITENS(i))-InstrRev(ARRAYITENS(i), "-", -1, 1))
					
					'verifica se pegou o total 
					if InstrRev(resto, ";", -1, 1) <> 0 then
						totalGeral = Cdbl(trim(right(resto,len(resto)-InstrRev(resto, ";", -1, 1))))
						resto	   = left(resto,InstrRev(resto, ";", -1, 1)-1)
					end if
									
					ARRAYQTD = Split(resto, "*", -1, 1)
					
					acuValor    = ""
					acuPeriodo  = ""
					acuValorTot = ""
					
					for x = 0 to UBOUND(ARRAYQTD)
						qtdPeriodo  = Cdbl(trim(right(ARRAYQTD(x),len(ARRAYQTD(x))-InstrRev(ARRAYQTD(x), ",", -1, 1))))
						periodo 	= trim(left(ARRAYQTD(x),InstrRev(ARRAYQTD(x), ",", -1, 1)-1))
						
						'verifica se o valor nao é zero
						if qtdPeriodo >0 then
							if len(acuValor) = 0 then
								acuValor 	= round((qtdPeriodo*100)/totalGeral)&""
								acuValorTot = round((qtdPeriodo*100)/totalGeral)&"% ("&qtdPeriodo&") "
							else
								acuValor 	= acuValor&","&round((qtdPeriodo*100)/totalGeral)&""
								acuValorTot = acuValorTot&" / "&round((qtdPeriodo*100)/totalGeral)&"% ("&qtdPeriodo&")"
							end if
						else
							if len(acuValor) = 0 then
								acuValor 	= "0"
								acuValorTot = "0% ("&qtdPeriodo&") "
							else
								acuValor 	= acuValor&", 0"
								acuValorTot = acuValorTot&" / 0% ("&qtdPeriodo&")"
							end if
						end if
						
						if len(acuPeriodo) = 0 then
							acuPeriodo = periodo
						else
							acuPeriodo = acuPeriodo&","&periodo
						end if
					next		
					
					'monta grafico automaticamente
					With .seriescollection.add
						.DataLabelsCollection.Add
						.caption = nome &" - " &acuValorTot &" Total = "&totalGeral
						
						'so mostra o periodo uma vez
						'if i=0 then
							.setdata c.chdimcategories,c.chdataliteral, acuPeriodo
						'end if
						
						.setdata c.chdimvalues,c.chdataliteral, acuValor
						'.Interior.Color = "green"
						
					End With
				
				'.Axes(c.chAxisPositionleft).NumberFormat = "#,##0"
				'.Axes(c.chAxisPositionleft).Scaling.Minimum = 0 
				'.Axes(c.chAxisPositionleft).Scaling.Maximum = 100
							
				next
				
			END IF		
		
		End With		
			
		.border.color = c.chColorNone

		'grava a imagem no caminho que o usuario
		Set arquivo = Server.CreateObject("Scripting.FileSystemObject")	
				
		'ADICIONA UM HORARIO PARA O CAMINHO
		caminho = LEFT(caminho, (len(caminho) - 4))&"_"&replace(formatahora(now), ":", "")&".gif"
	
		.ExportPicture caminho, "gif", largura, altura
					
		geraGrafico = right(caminho,len(caminho)-InstrRev(caminho, "\", -1, 1))
			
	End With
	
end function


'FUNCAO QUE FAZ A VALIDACAO
'RETORNO = 0 CERTO | 1 ERRADO
function CalculaCNPJ(a)

	Dim RecebeCNPJ, Numero(14), soma, resultado1, resultado2
	
	RETORNO = 0
	RecebeCNPJ = a

	s="" 
	
	for x =1  to len(RecebeCNPJ)
		ch = mid(RecebeCNPJ,x,1)
		if asc(ch)>=48 and asc(ch)<=57 then
			s=s & ch
		end if
	next
	
	RecebeCNPJ = s

	if len(RecebeCNPJ) <> 14 then
		RETORNO = 1
	elseif RecebeCNPJ = "00000000000000" then
		RETORNO = 1
	else

		Numero(1) = Cint(Mid(RecebeCNPJ,1,1))
		Numero(2) = Cint(Mid(RecebeCNPJ,2,1))
		Numero(3) = Cint(Mid(RecebeCNPJ,3,1))
		Numero(4) = Cint(Mid(RecebeCNPJ,4,1))
		Numero(5) = Cint(Mid(RecebeCNPJ,5,1))
		Numero(6) = CInt(Mid(RecebeCNPJ,6,1))
		Numero(7) = Cint(Mid(RecebeCNPJ,7,1))
		Numero(8) = Cint(Mid(RecebeCNPJ,8,1))
		Numero(9) = Cint(Mid(RecebeCNPJ,9,1))
		Numero(10) = Cint(Mid(RecebeCNPJ,10,1))
		Numero(11) = Cint(Mid(RecebeCNPJ,11,1))
		Numero(12) = Cint(Mid(RecebeCNPJ,12,1))
		Numero(13) = Cint(Mid(RecebeCNPJ,13,1))
		Numero(14) = Cint(Mid(RecebeCNPJ,14,1))
		
		soma = Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2
		
		soma = soma -(11 * (int(soma / 11)))
		
		if soma = 0 or soma = 1 then
			resultado1 = 0
		else
			resultado1 = 11 - soma
		end if
		
		if resultado1 = Numero(13) then
			soma = Numero(1) * 6 + Numero(2) * 5 + Numero(3) * 4 + Numero(4) * 3 + Numero(5) * 2 + Numero(6) * 9 + Numero(7) * 8 + Numero(8) * 7 + Numero(9) * 6 + Numero(10) * 5 + Numero(11) * 4 + Numero(12) * 3 + Numero(13) * 2
			soma = soma - (11 * (int(soma/11)))
			
			if soma = 0 or soma = 1 then
				resultado2 = 0
			else
				resultado2 = 11 - soma
			end if
		
			if resultado2 = Numero(14) then
				RETORNO = 0
			else
				RETORNO = 1
			end if
		else
			RETORNO = 1
		end if
	end if			
	
	CalculaCNPJ = RETORNO
	
end function
 
'FUNCAO QUE FAZ A VALIDACAO
'RETORNO = 0 CERTO | 1 ERRADO
Function verificaEmail(StrMail)
 
	' Função que verifica validação de preenchimento de E-Mail.	
	' Se há espaço vazio, então...	 
	If InStr(1, StrMail, " ") > 0 Then
	 
		verificaEmail = 1	 
		Exit Function
	 
	End If	
	 
	' Verifica tamanho da String, pois o menor endereço válido é x@x.xx.
	If Len(FncStrSpace(StrMail)) < 6 Then
	 
		verificaEmail = 1	 
		Exit Function
	 
	End If
		 
	' Verifica se se há um "@" no endereço.	 
	If InStr(FncStrSpace(StrMail), "@") = 0 Then
	 
		verificaEmail = 1	 
		Exit Function
	 
	End If
	 
	' Verifica se há um "." no endereço.	 
	If InStr(FncStrSpace(StrMail), ".") = 0 Then
	 
		verificaEmail = 1	 
		Exit Function
	 
	End If
		 
	' Verifica se há a quantidade mínima de caracteres é igual ou maior que 3.	 
	If Len(FncStrSpace(StrMail)) - InStrRev(FncStrSpace(StrMail), ".") > 3 Then
	 
		verificaEmail = 1	 
		Exit Function
	 
	End If	
	
	' Verifica se há "_" após o "@".	 
	If InStr(FncStrSpace(StrMail), "_") <> 0 And InStrRev(StrMail, "_") > InStrRev(FncStrSpace(StrMail), "@") Then
	 
		verificaEmail = 1	 
		Exit Function
	 
	Else
	 
		Dim IntCounter	 
		Dim IntF
	 
		IntCounter = 0
	 
		For IntF = 1 To Len(FncStrSpace(StrMail))
	 
			If Mid(StrMail, IntF, 1) = "@" Then	 
				IntCounter = IntCounter + 1	 
			End If
	 
		Next
	 
		If IntCounter > 1 Then
	 
			verificaEmail = 0
	 
		End If		
		
	 
		' Valida cada caracter do endereço.	 
		IntF = 0
	 
		For IntF = 1 To Len(FncStrSpace(StrMail))
	 
			If IsNumeric(Mid(FncStrSpace(StrMail), IntF, 1)) = False And _
	 
				(LCase(Mid(FncStrSpace(StrMail), IntF, 1)) < "a" Or _
	 
				LCase(Mid(FncStrSpace(StrMail), IntF, 1)) > "z") And _
	 
				Mid(FncStrSpace(StrMail), IntF, 1) <> "_" And _
	 
				Mid(FncStrSpace(StrMail), IntF, 1) <> "." And _
	 
				Mid(FncStrSpace(StrMail), IntF, 1) <> "-" Then
	 
				verificaEmail = 0
	 
			End If
	 
		Next
			 
	End If
 
End Function
 
Function FncStrSpace(StrAddress)
 
	' Reduz os espaços em branco.	 
	FncStrSpace = Trim(LTrim(RTrim(StrAddress)))
 
End Function

'-----------------------------------------------------
Function EnviaEmail(p_strNomeFROM, p_strEmailFROM, p_strNomeTO, p_strEmailTO, p_strEmailCC, p_strEmailBCC, p_strTipoEmail, p_strAssuntoEmail, p_strMensagem, p_strAnexo)

    'Contem o SMTP padrao para o Envio de Emails
'Const strSMTP = "192.168.10.100"

'-----------------------------------------------------
'Funcao: EnviaEmailJOVE(p_strNomeFROM, p_strEmailFROM, p_strNomeTO, p_strEmailTO, p_strEmailCC, p_strEmailBCC, p_strTipoEmail, p_strAssuntoEmail, p_strMensagem)
'Sinopse: Envia um email utilizando o componente configurado
'Parametros:
'            p_strNomeFROM          : Contem o Nome do Remetente
'            p_strEmailFROM         : Contem o Email do Remetente
'            p_strNomeTO            : Contem o Nome do Destinatario
'            p_strEmailTO           : Contem o Email ou Emails de quem vai receber a mensagem
'            p_strEmailCC           : Contem o Email ou Emails de quem vai receber a copia da mensagem
'            p_strEmailBCC          : Contem o Email ou Emails de quem vai receber a copia oculta da mensagem
'            p_strTipoEmail         : Tipo de Mensagem que sera enviada ("HTML" ou "TEXTO")
'            p_strAssuntoEmail      : Contem o Assunto do Email
'            p_strMensagem          : Mensagem do Email
'Retorno: String ("OK" quando for executada com sucesso)
'Autor: Gabriel Fróes (www.codigofonte.com.br)

Const strComponenteEmail = "CDOSYS"  	
	'****************************************************
	
	If strComponenteEmail = "CDOSYS" Then
	
		sch = "http://schemas.microsoft.com/cdo/configuration/"
		Set cdoConfig = Server.CreateObject("CDO.Configuration")
		
		'Variaveis
		
		Dim meuservidorsmtp
		Dim minhacontaautenticada
		Dim minhasenhaparaenvio
		Dim emailorigem
		Dim emaildestino
		
		'Abaixo seguem algumas definicoes de variaveis para o envio de seu formulario. Por favor preencha os campos abaixo.
		meuservidorsmtp = "smtp.gmail.com" ' Informacoes so seu servidor SMTP		
		minhacontaautenticada = "aangeladosanjos@gmail.com" ' conta de e-mail utilizada para enviar
		minhasenhaparaenvio = "potinha1702" ' senha da conta de e-mail
		
		'Fim da definição manual de parâmetros.
		
		'Fim da definição manual de parâmetros.		
		cdoConfig.Fields.Item(sch & "sendusing") = 2
		cdoConfig.Fields.Item(sch & "smtpauthenticate") = 1
		cdoConfig.Fields.Item(sch & "smtpusessl") = True 
		cdoConfig.Fields.Item(sch & "smtpserver") = meuservidorsmtp
		cdoConfig.Fields.Item(sch & "smtpserverport") = 465
		cdoConfig.Fields.Item(sch & "smtpconnectiontimeout") = 30
		cdoConfig.Fields.Item(sch & "sendusername") = minhacontaautenticada
		cdoConfig.Fields.Item(sch & "sendpassword") = minhasenhaparaenvio
		cdoConfig.fields.update		
		
		Set cdoMessage = Server.CreateObject("CDO.Message")
		Set cdoMessage.Configuration = cdoConfig
		
		cdoMessage.BodyPart.Charset = "iso-8859-1"
		cdoMessage.From 			= p_strEmailFROM
		cdoMessage.To 				= p_strEmailTO
		cdoMessage.CC        		= p_strEmailCC
		cdoMessage.BCC        		= p_strEmailBCC
		cdoMessage.Subject 			= p_strAssuntoEmail


        'Verificando o tipo de mensagem (default é TEXTO)
        If p_strTipoEmail = "HTML" Then
            cdoMessage.HTMLBody    = p_strMensagem
        Else
            cdoMessage.TextBody    = p_strMensagem
        End If
		
		'ANEXO
		IF LEN(p_strAnexo)>0 THEN
			cdoMessage.AddAttachment p_strAnexo
		END IF
		
		on error resume next

		cdoMessage.Send
			
		If Err.Number = 0 Then
            EnviaEmail = 0
        Else
			EnviaEmail = 1
            'EnviaEmailJOVE = "Não foi possível enviar o email, por favor entre em contato com <a href='mailto:" & p_strEmailTO & "?subject=Erro%20ao%20enviar%20email%20pelo%20site'>" & p_strEmailTO & "</a> e escreva o erro abaixo."
        End If
		
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing
		
		If Err.Number <> 0 Then
			EnviaEmail = 1
            'EnviaEmailJOVE = EnviaEmailJOVE & "<br>Erro:" & Err.Description
        End If
        Exit Function ' N EXECUTA A PROXIMA FUNCAO 

		
	End If 'Fim da Utilização do CDOSYS	   
   
   
End Function%>
