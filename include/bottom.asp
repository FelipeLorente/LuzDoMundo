
<style>
.pagination > li > span:hover {
    z-index: 3;
    color: #000000;
    background-color: #ffffff;
}
.pagination > li.active a {
	z-index: 3;
	color: #ffffff;
	font-weight:bold;
    background-color: #003d7b;
	border-color: #003d7b;
}

.pagination > li a {
	color: #000000;
    background-color: #ffffff;
}

.pagination > li a:hover {
	color: #ffffff;
    background-color: #003d7b;
}

</style>

<%'CRIA TABELA PARA MOSTRAR O NUMERO DE PAG
SUB LinksNavegacao()%>

    <div class="text-center">
    	<nav>
            <ul class="pagination justify-content-end " style="margin:0px;">
            	<li>
            		<%IF PagAtual > 1 THEN%>
            			<a onClick="BuscaReg('Anterior')" aria-label="Voltar Página" style="cursor:pointer">
            				<span aria-hidden="true" style="font-weight:bold">&laquo; </span>
           				</a>
               		<%ELSE%>
            			<span><img src="../imagens/bullet.gif" width="5" height="20"></span>
            		<%END IF%> 
            	</li>
                
            	<%for i = qtdmax - (qtdpag - 1) to qtdmax
            	IF i <= TotPag THEN
                	IF i <> CInt(PagAtual) THEN%>
                    	<li>
                        	<a onClick="BuscaReg('<%=i%>')" style="cursor:pointer"><%=i%></a>
                        </li>
                	<%ELSE%>
						<%IF PagAtual <> TotPag+1 THEN%>
                            <li class="active"><a  style="height:34px"><%=i%></a></li>
                        <%END IF
                	END IF
            	END IF
            	Next%>
                
            	<li>
					<%IF PagAtual <> TotPag THEN%>
                    	<a onClick="BuscaReg('Proxima')" aria-label="Próxima Página" style="cursor:pointer">
                    		<span aria-hidden="true" style="font-weight:bold">&raquo;</span>
            			</a>
            		<%ELSE%>
            			<span><img src="../imagens/bullet.gif" width="5" height="20"></span>
            		<%END IF%>
            	</li>
            </ul>
    	</nav>                      
    </div> 
    
    <script>
    function BuscaReg(v){
		$('#forme').attr('action','<%=URL%>?PagAtual=<%=PagAtual%>&qtdpag=<%=qtdpag%>&qtdmax=<%=qtdmax%>&Submit='+v+'&paginaDescPos=<%=paginaDescPos%>&paginaDesc=<%=paginaDesc%>');
		$('#forme').submit();
	}
    </script>
    
    
<%END SUB%>