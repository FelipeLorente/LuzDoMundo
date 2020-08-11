<%'PARAMETROS OBRIGATORIO
submit 			= Request.QueryString("Submit")
qtdmax 			= Request.QueryString("qtdmax")
PagAtual 		= Request.QueryString("PagAtual")


'--------------------- PAGINAÇÃO ------------------------------

'Número total de registros a serem exibidos por página
Const RegPorPag = 5

'Número de páginas a ser exibido no índice de paginação
qtdpag = 10%>

<%'PAGINAÇÃO
IF LEN(pagAtual) 	= 0 THEN
  PagAtual 			= 1
  qtdmax 			= qtdpag
ELSE
  qtdmax 			= CInt(qtdmax)
  PagAtual 			= CInt(pagAtual)
  
  'SELECIONA A PAG DE ACORDO COM submit
  SELECT CASE submit
    CASE "Anterior" 	: PagAtual 	= PagAtual - 1
    CASE "Proxima" 		: PagAtual 	= PagAtual + 1
    CASE "Menos" 		: qtdmax 	= qtdmax - qtdpag
    CASE "Mais" 		: qtdmax 	= qtdmax + qtdpag
    CASE ELSE 			: PagAtual 	= CInt(submit)
  END SELECT
  
  IF qtdmax < PagAtual THEN
    qtdmax = qtdmax + qtdpag
  END IF
  IF qtdmax - (qtdpag - 1) > PagAtual THEN
    qtdmax = qtdmax - qtdpag
  END IF
END IF%>

<html>
  <head>  

    <title>CADASTRO CLOUD - JOVE</title>
	<!--<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />-->
	<meta name="viewport" content="width=device-width, initial-scale=1.0,charset=iso-8859-1">
    
    <!--imports do select flex-->
    <link href="css/SelectFlex/flexselect.css" rel="stylesheet" type="text/css">
    <script src="js/SelectFlex/javascriptSelectFlex.js"></script>
    <script src="js/SelectFlex/jquery.flexselect.js"></script>
          	    
    <!-- Bootstrap core CSS -->    
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap.css" />
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap.min.css" />
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap-theme.css" />    
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap-theme.min.css" />   
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap-select.css">   
    
    <script language="javascript" src="BootStrap/js/npm.js"></script>      
    <script language="javascript" src="BootStrap/js/jquery.min.js"></script>     
    <script language="javascript" src="BootStrap/js/jquery.validate.min.js"></script>     
    <script language="javascript" src="BootStrap/js/bootstrap.js"></script>   
    <script language="javascript" src="BootStrap/js/bootstrap.min.js"></script>   
    <script language="javascript" src="BootStrap/js/js.js" type="text/javascript"></script> 
    <script language="javascript" src="BootStrap/js/bootstrap-select.js"></script>   
    <script language="javascript" src="BootStrap/js/jquery.mask.min.js"></script>  
    
    <style>
		body {
		  padding-top: 0px;
		}
		/*.starter-template {
		  padding: 40px 15px;
		  text-align: center;
		}*/
		
	  </style>
  
  </head>
  

  <body topmargin="0" leftmargin="0">
  
<!-- MODAL AVISO glyphicon-remove-->
<div class="modal fade" id="aviso" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
  <div class="modal-dialog" id="avisoClass" role="document">
    <div class="modal-content">
      <div class="modal-header" style="background:#003d7b">
        <h4 class="modal-title" id="avisoTitulo" style="color:#FFFFFF;font-weight:bold;font-size:16px"></h4>
      </div>
      <div class="modal-body">
      		<table border="0">
            <tr> 
                <td style="padding-right:30px;" width="10%"> 
                    <i id="avisoImagem" class="glyphicon" style="font-size:36px"></i>
                </td>
                <td width="80%" id="avisoDescricao"></td>
            </tr>
            </table>
      </div>
      <div class="modal-footer">
         <input type="button" id="ok" name="ok" value="Ok" class="btn btn-primary" data-dismiss="modal" aria-label="Close" style="background:#003d7b;">     
      </div>
    </div>
  </div>
</div>

<!-- MODAL CONFIRMA glyphicon-remove-->
<div class="modal fade" id="avisoConfirma" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
  <div class="modal-dialog" id="avisoClassConfirma" role="document">
    <div class="modal-content">
      <div class="modal-header" style="background:#003d7b">
        <h4 class="modal-title" id="avisoTituloConfirma" style="color:#FFFFFF;font-weight:bold;font-size:16px"></h4>
      </div>
      <div class="modal-body">
      		<table border="0">
            <tr> 
                <td style="padding-right:30px;" width="10%"> 
                    <i id="avisoImagemConfirma" class="glyphicon" style="font-size:36px"></i>
                </td>
                <td width="80%" id="avisoDescricaoConfirma"></td>
            </tr>
            </table>
      </div>
      <div class="modal-footer">
      	 <input type="button" id="nao" name="nao" value="Não" class="btn btn-primary" data-dismiss="modal" aria-label="Close" style="background:#003d7b;">
         <input type="button" id="sim" name="sim" value="Sim" class="btn btn-primary" data-dismiss="modal" style="background:#003d7b;">     
      </div>
    </div>
  </div>
</div>

<!-- MODAL CARREGANDO REGISTRO-->
<div class="modal fade" id="carregandoReg" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" style="
    position: fixed;
    width: 100%;
    height: 100%;
    left: 0;
    top: 0;
    background: rgba(100,100,100,0.3);
    background-image: url(imagens/carregandoRed.gif);
    background-repeat: no-repeat;
    background-position: center center;
    background-size: 60px;
  ">
  
</div>

<!--progresso da barra-->
<div class="modal fade" id="progressoBar">         
	<div class="modal-dialog" id="avisoClassConfirma" role="document">
        <div class="progress">
            <div id="progressoBarPor" class="progress-bar progress-bar-danger" role="progressbar" aria-valuenow="10" aria-valuemin="0" aria-valuemax="100">
                <b id="progressoBarPorTexto"></b> 
            </div>                                
        </div>
    </div>
</div>

