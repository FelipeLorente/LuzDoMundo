<%
'REDIRECIONA PAGINA QUANDO EXPIRA
FUNCTION REDIRECTPAGE(TIPOPAG)
	
	'1=REDIRECIONA PAGINA PRINCIPAL
	IF TIPOPAG = 1 THEN%>
		<script>
			location = 'default.asp?erro=2'
		</script>
	<%ELSEIF TIPOPAG = 2 THEN%>
		<script>
			parent.location = 'default.asp?erro=2'
		</script>
    <%ELSEIF TIPOPAG = 3 THEN%>
		<script>
			parent.parent.location = 'default.asp?erro=2'
		</script>    
	<%END IF
	
END FUNCTION
%>