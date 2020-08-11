// validador de qualquer formulário 
function validaform(obj){
var form = obj;

	/* loop em todos os campos do formulário */
	for (x = 0; x < form.elements.length; x++){
	
		/* se o campo estiver desabilitado ou apenas leitura não verifica */	
		if (form.elements[x].disabled == false || form.elements[x].readonly == false){
			
			/* valida text, password, select, select-one ou file */
			if (form.elements[x].type == 'text' || form.elements[x].type == 'textarea' ||  form.elements[x].type == 'password' || form.elements[x].type == 'select-one' || form.elements[x].type == 'file'){		
			
				//acha o campo email
				if(form.elements[x].getAttribute("parametro") == 'email'){					
					//verifica se existe valor para validar
					if (form.elements[x].value.length > 0){
						//valida email com (@) e (.)
						if (form.elements[x].value.indexOf("@")==-1 || form.elements[x].value.indexOf(".")==-1){ 
													
							// chama função do alert
							chamaAlerta('Digite um email válido!', form.elements[x], '')							
							return false
						}
					}
					
				}
				
				//acha o campo data
				if(form.elements[x].getAttribute("parametro") == 'data'){
					
					//verifica se existe valor para validar
					if (form.elements[x].value.length > 0){
						barra(form.elements[x])
							
						if(form.elements[x].value.length != 10){
																		
							// chama função do alert
							chamaAlerta("Digite o data no formato 'dd/mm/aaaa!'", form.elements[x], '')																
							form.elements[x].value = ""	
							return false
						}
					}
				}			
				
				/* se for obrigatório e aida não estiver preeenchido */
				if (form.elements[x].getAttribute("requerido") == 'yes' && form.elements[x].value.trim().length == 0){
					var msg = form.elements[x].getAttribute("msg");
					form.elements[x].value = ''
					
					if (msg == null){
						var msg = 'Campo obrigatório. ('+ form.elements[x].name +')'  
					}			
					
					// chama função do alert
					chamaAlerta(msg, form.elements[x], '')						

					return false;
				}

			}			
		}	
	}

	return true;
}

// verifica se é numero, se nao for. retira do campo
var verifNum = function(valor, campo) {	
	
	// faz replace do valor
	valor = valor.replace(' ' , '').replace('-' , '').replace('/' , '').replace('/' , '').replace('.' , '')

	// verifica se tem algo digitado
	if (valor.length > 0){
			
		// verifica se é numero	
		if(isNaN(valor)){
			$('#'+campo).val('')	
			
			chamaAlerta('Este campo só aceita números!', '#'+campo, '')			
		}
	}
}

//função que chama o alert da pagina topo
var chamaAlerta = function(msg, elemento, url) {

	// mostra aviso de erro
	$('#aviso').modal('show')
	$('#avisoTitulo').text('Aviso do Sistema')
	$('#avisoDescricao').text(msg)
	$('#avisoImagem').addClass('glyphicon-remove text-danger');
	
	// foco no botao
	$('#aviso').on('shown.bs.modal', function(){
	  $('#ok').focus();			  
	});	
	
	// foco no campo apos sair
	$('#aviso').on('hidden.bs.modal', function () {
		
		if(url == ''){
			// do something…
			elemento.focus();	
		}else{
			$(location).attr('href',url)
		}
	})	
}

//função que chama o alert da pagina topo
function chamaConfirma(msg, url){

	// mostra aviso de erro
	$('#avisoConfirma').modal('show')
	$('#avisoTituloConfirma').text('Aviso do Sistema')	
	$('#avisoImagemConfirma').removeClass('glyphicon-remove text-danger')	
	$('#avisoImagemConfirma').addClass('glyphicon-warning-sign text-danger')
	$('#avisoDescricaoConfirma').text(msg)
	
	// foco no botao
	$('#avisoConfirma').on('shown.bs.modal', function(){
	  $('#nao').focus();			 	  
	});
	
	// se sim excluir registros
	$('#sim').click(function(){		
		$(location).attr('href',url);		
	})	
}

//onkeyup="barra(this)"
function barra(campo){
	//verifica se for >10 zera o campo data
	if(campo.value.length > 10){
		campo.value = ''
	}
	
	//coloca as (/) no campo data
	if(campo.value.length == 2){
		campo.value = campo.value +"/"
	}
	
	if(campo.value.length == 5){
		campo.value = campo.value +"/"
	}
	
	//funcao para validacao de data erradas
	if(campo.value.length == 10){
		
		//cria variaveis para separar a data
		dia = campo.value;
		mes = campo.value;
		ano = campo.value;	

		//variaveis de apoio passando status
		var diaCerto = ''
		var mesCerto = ''
		var anoCerto = ''
				
		//separa a data em dd mm aaaa e verifica inconsistencias
		dia = dia.substring (0, 2)
		mes = mes.substring (3, 5)
		ano = ano.substring (6, 10)
		
		//verifica se os parametros da data esta correto
		if(dia.trim() > 0 && dia.trim() < 32){			
			diaCerto = "yes"			
		}else{
			diaCerto = "no"	
		}		
		if(mes.trim() > 0 && mes.trim() < 13){
			mesCerto = "yes"	
		}else{
			mesCerto = "no"	
		}
		if(ano.trim().length != 4){
			anoCerto = "no"				
		}else{
			anoCerto = "yes"	
		}

		//verifica se os status estao ok.
		if(diaCerto == "no" || mesCerto == "no" || anoCerto == "no"){
			
			// chama função do alert
			chamaAlerta('Data incorreta!', campo, '')	
			campo.value = ""		
			return false
		}			
	}
}

//onKeypress="return somenteNumeros(event);"
//SOMENTE NUMEROS
function somenteNumeros(obj){
	var charCode = (obj.which) ? obj.which : event.keyCode;
	return (charCode>=48 && charCode<=57 || charCode<20);
}

//onKeyUp="mascaraCep(this)"
//coloca mascara para cep
function mascaraCep(obj){
	//valida campo cep colocando traço após 5 numeros
	if(obj.value.length == 5){
		obj.value = obj.value+ "-"
	}
}

//onblur="mascaraContato(this)"
//mascara pra TEL/CEL quando sai do campo
function mascaraContato(obj){
	//quando for 8 digitos é tel
	if(obj.value.length == 8){
		obj.value = obj.value.substring(0,4)+"-"+obj.value.substring(4,8);
	}	
	else{
		//senão é celular e padroniza o traço (-)
		obj.value = obj.value.substring(0,5)+"-"+obj.value.substring(5,9);
	}	
}

//onKeyUp="mascaraTel(this)"
//coloca mascara para tel
function mascaraTel(obj){
	if(obj.value.length == 2){
		obj.value = obj.value+ " "
	}
	if(obj.value.length == 7){
		obj.value = obj.value+ "-"
	}
}

//onKeyUp="mascaraTel(this)"
//coloca mascara para tel
function mascaraCel2(obj){
	if(obj.value.length == 2){
		obj.value = obj.value+ " "
	}
	if(obj.value.length == 8){
		obj.value = obj.value+ "-"
	}
}

//onKeyUp="mascaraCel(this)"
//coloca mascara para cel
function mascaraCel(obj){
	//valida campo tel colocando traço após 4 numeros
	if(obj.value.length == 5){
		obj.value = obj.value+ "-"
	}
}
//onKeyUp="mascaraCnpj(this)"
function mascaraCnpj(obj){
//valida campo tel colocando traço após 4 numeros
	if(obj.value.length == 2){
		obj.value = obj.value+ "."
	}
	if(obj.value.length == 6){
		obj.value = obj.value+ "."
	}
	if(obj.value.length == 10){
		obj.value = obj.value+ "/"
	}
	if(obj.value.length == 15){
		obj.value = obj.value+ "-"
	}
}

//onKeyUp="mascaraCpf(this)"
function mascaraCpf(obj){
//valida campo tel colocando traço após 4 numeros
	if(obj.value.length == 3){
		obj.value = obj.value+ "."
	}
	if(obj.value.length == 7){
		obj.value = obj.value+ "."
	}
	if(obj.value.length == 11){
		obj.value = obj.value+ "-"
	}
}

// apaga registro
function deletar(obj){
	if(confirm("Deseja realmente apagar este item?")){
		obj.tipo.value = 3
		obj.submit()
	}
}

//-----------------------------------------------------
//Funcao: MascaraMoeda
//Sinopse: Mascara de preenchimento de moeda
//Parametro:
//   objTextBox : Objeto (TextBox)
//   SeparadorMilesimo : Caracter separador de milésimos
//   SeparadorDecimal : Caracter separador de decimais
//   e : Evento
//Retorno: Booleano
//Autor: Gabriel Fróes - www.codigofonte.com.br
//-----------------------------------------------------
function MascaraMoeda(objTextBox, SeparadorMilesimo, SeparadorDecimal, e){
	
    var sep = 0;
    var key = '';
    var i = j = 0;
    var len = len2 = 0;
    var strCheck = '0123456789';
    var aux = aux2 = '';
    var whichCode = (window.Event) ? e.which : e.keyCode;
    if (whichCode == 13) return true;
    key = String.fromCharCode(whichCode); // Valor para o código da Chave
    if (strCheck.indexOf(key) == -1) return false; // Chave inválida
    len = objTextBox.value.length;
    for(i = 0; i < len; i++)
        if ((objTextBox.value.charAt(i) != '0') && (objTextBox.value.charAt(i) != SeparadorDecimal)) break;
    aux = '';
    for(; i < len; i++)
        if (strCheck.indexOf(objTextBox.value.charAt(i))!=-1) aux += objTextBox.value.charAt(i);
    aux += key;
    len = aux.length;
    if (len == 0) objTextBox.value = '';
    if (len == 1) objTextBox.value = '0'+ SeparadorDecimal + '0' + aux;
    if (len == 2) objTextBox.value = '0'+ SeparadorDecimal + aux;
    if (len > 2) {
        aux2 = '';
        for (j = 0, i = len - 3; i >= 0; i--) {
            if (j == 3) {
                aux2 += SeparadorMilesimo;
                j = 0;
            }
            aux2 += aux.charAt(i);
            j++;
        }
        objTextBox.value = '';
        len2 = aux2.length;
        for (i = len2 - 1; i >= 0; i--)
        objTextBox.value += aux2.charAt(i);
        objTextBox.value += SeparadorDecimal + aux.substr(len - 2, len);
    }
    return false;
}