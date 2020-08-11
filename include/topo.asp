<html>
  <head>  

    <title>LUZ DO MUNDO</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    
    <!--imports do select flex-->
    <link href="css/SelectFlex/flexselect.css" rel="stylesheet" type="text/css">
    <script src="js/SelectFlex/javascriptSelectFlex.js"></script>
    <script src="js/SelectFlex/jquery.flexselect.js"></script>
    
    <!-- CSS programado-->          	    
	<link rel="stylesheet" type="text/css" href="BootStrap/css/style.css" />
                    
    <!-- Bootstrap core CSS -->    
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap.css" />
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap.min.css" />
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap-theme.css" />    
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap-theme.min.css" />   
    <link rel="stylesheet" type="text/css" href="BootStrap/css/bootstrap-select.css">   
    <link rel="stylesheet" type="text/css" href="Bootstrap/css/bootstrap-select.min.css">
    <link rel="stylesheet" type="text/css" href="Bootstrap/css/bootstrap-datetimepicker.css" /> 
    <link rel="stylesheet" type="text/css" href="Bootstrap/css/font-awesome.min.css">  
    
    <script language="javascript" src="BootStrap/js/npm.js"></script>      
    <script language="javascript" src="BootStrap/js/jquery.min.js"></script>     
    <script language="javascript" src="BootStrap/js/jquery.validate.min.js"></script>     
    <script language="javascript" src="BootStrap/js/bootstrap.js"></script>   
    <script language="javascript" src="BootStrap/js/bootstrap.min.js"></script>   
    <script language="javascript" src="BootStrap/js/js.js" type="text/javascript"></script> 
    <script language="javascript" src="BootStrap/js/bootstrap-select.js"></script> 
    <script language="javascript" src="Bootstrap/js/bootstrap-select.min.js"></script>
    <script language="javascript" src="Bootstrap/js/moment-with-locales.js"></script>
    <script language="javascript" src="Bootstrap/js/bootstrap-datetimepicker.js"></script>      
    
    <style>
		body {
		  padding-top: 80px;		  
		}
		.starter-template {
		  padding: 20px 15px;
		  text-align: center;
		}
		
	  </style>
  </head>

  <body>
  
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


