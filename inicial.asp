<!--#include file="include/conexao.asp"-->
<!--#include file="include/topo.asp"-->
<!--#include file="include/expiraSessao.asp"-->
  
<%'VERIFICA A SESSION
IF LEN(SESSION("usu_cod")) = 0 THEN
	REDIRECTPAGE(1)	
	response.End() 
END IF%>

<style>

.navbar-default .navbar-nav > .open > a:focus .caret {
  border-top-color: #3B5998 !important;
  border-bottom-color: #3B5998 !important;
}

</style>

<nav class="navbar navbar-inverse navbar-fixed-top" style="background:#003d7b">
    <div class="container">
        <div class="navbar-header">        
        	<font class="navbar-brand" style="color:#FFFFFF; font-weight:bold">Luz do Mundo</font> 
        </div>    	
        <ul class="nav navbar-nav navbar-right">        
            <li><a href="sair.asp"><span class="glyphicon glyphicon-log-out"></span> Sair</a></li>
        </ul>
    </div>
</nav>

<div class="container" style="width:90%; height:95%;">  
    <ul class="list-inline nav nav-tabs" id="menuGeral">
        <li id="menuPai0"><a name="link0" id="link0" data-toggle="tab" href="#menuFilho0" onClick="mudaMenu(this)">Igrejas </a></li>
        <li id="menuPai1"><a name="link1" id="link1" data-toggle="tab" href="#menuFilho1" onClick="mudaMenu(this)">Irmãos </a></li>
        <li id="menuPai3"><a name="link3" id="link3" data-toggle="tab" href="#menuFilho3" onClick="mudaMenu(this)">Aniversários </a></li>
        <li id="menuPai2"><a name="link2" id="link2" data-toggle="tab" href="#menuFilho2" onClick="mudaMenu(this)">Usuários </a></li>
    </ul>    
    <div class="tab-content">
        <div id="menuFilho0" class="tab-pane">     
          <iframe name="igrejas" id="igrejas" width="100%" height="90%" frameborder="0" allowfullscreen="allowFullScreen"></iframe>
        </div>
        <div id="menuFilho1" class="tab-pane">     
          <iframe name="irmaos" id="irmaos" width="100%" height="90%" frameborder="0" allowfullscreen="allowFullScreen"></iframe>
        </div> 
        <div id="menuFilho2" class="tab-pane">     
          <iframe name="usuarios" id="usuarios" width="100%" height="90%" frameborder="0" allowfullscreen="allowFullScreen"></iframe>
        </div>
        <div id="menuFilho3" class="tab-pane">     
          <iframe name="aniversarios" id="aniversarios" width="100%" height="90%" frameborder="0" allowfullscreen="allowFullScreen"></iframe>
        </div>                 
    </div>
</div>

<!--faixa azul do rodape-->
<nav class="navbar navbar-inverse navbar-fixed-bottom" style="background:#003d7b; min-height:20px;max-height:20px; text-align:center;">
	<label style="color:#fff; font-size:11px; padding-top:2px;"><%=ucase(periodo & " " & SESSION("usu_nome") & ", HOJE É DIA " & MontaDataExtenso())%></label>
</nav>
</body>
</html>

<script>	
// bloqueia menu
function blockMenu(){
	
	$('#menuPai0').addClass('disabled');
	$('#menuPai1').addClass('disabled');
	$('#menuPai2').addClass('disabled');
	$('#menuPai3').addClass('disabled');
	
	$('#link0').removeAttr("data-toggle");
	$('#link1').removeAttr("data-toggle");
	$('#link2').removeAttr("data-toggle");
	$('#link3').removeAttr("data-toggle");
	
	$('#link0').attr("onClick","");
	$('#link1').attr("onClick","");
	$('#link2').attr("onClick","");
	$('#link3').attr("onClick","");
}

// desbloqueia menu
function desblockMenu(){
	
	$('#menuPai0').removeClass('disabled');
	$('#menuPai1').removeClass('disabled');
	$('#menuPai2').removeClass('disabled');
	$('#menuPai3').removeClass('disabled');
	
	$('#link0').attr("data-toggle", "tab");
	$('#link1').attr("data-toggle", "tab");
	$('#link2').attr("data-toggle", "tab");
	$('#link3').attr("data-toggle", "tab");
	
	$('#link0').attr("onClick","mudaMenu(this)");
	$('#link1').attr("onClick","mudaMenu(this)");
	$('#link2').attr("onClick","mudaMenu(this)");
	$('#link3').attr("onClick","mudaMenu(this)");
}

// seleciona o menu pelo click
function mudaMenu(obj, valor){
	var nome = obj.id;

	if(nome == 'link0'){
		$('#igrejas').attr('src','igrejas.asp')
	} 
	
	if(nome == 'link1'){
		$('#irmaos').attr('src','irmaos.asp?irm_cod=<%=valor%>')
	} 
	
	if(nome == 'link2'){
		$('#usuarios').attr('src','usuarios.asp')
	} 
	
	if(nome == 'link3'){
		$('#aniversarios').attr('src','aniversarios.asp')
	} 
}

//chama função quando a pag estiver carregada
$(document).ready(function() {
	$('#link0').trigger('click');	
})
</script>


<!--#include file="include/fechaConexao.asp"-->