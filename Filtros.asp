<!--#include file="Conexion.asp"-->
<%
	rep=Request.QueryString("rep")
	
	eeff=Request.QueryString("eeff")
	sup=Request.QueryString("sup")
	anio=Request.QueryString("anio")
	per=Request.QueryString("per")
	sec=Request.QueryString("sec")
	tip=Request.QueryString("tip")
	met=Request.QueryString("met")
	letr=Request.QueryString("letr")
	niv=Request.QueryString("niv")
	SQL = ""

	tabla ="SMVMA_BALGEN"
	if eeff="01" then
		tabla="SMVMA_BALGEN"
	elseif eeff="02" then
		tabla="SMVMA_ESTGP"
	elseif eeff="03" then
		tabla="SMVMA_FLEFEC"
	elseif eeff="04" then
		tabla="SMVMA_CAMPAT"
	end if

	Select Case rep
	Case "anio_Sup"
		SQL = "SELECT distinct ANIO_INF as cod, ANIO_INF as des FROM SMVMA_DIRECTORIO where Supervision='"& sup &"' order by 1 desc"
	Case "periodo_SupAnio"

		filt =" and (BalGen = 1 or GanPer =1 or FluEf =1 or CamPat =1)"
		if eeff="01" then
			filt=" and BalGen = 1"
		elseif eeff="02" then
			filt=" and GanPer =1"
		elseif eeff="03" then
			filt=" and FluEf =1"
		elseif eeff="04" then
			filt=" and CamPat =1"
		end if

		SQL = "SELECT distinct PERINF_EMP as cod, case PERINF_EMP when 'A' then 'ANUAL' else 'TRIM '+PERINF_EMP end as des FROM SMVMA_DIRECTORIO where Supervision='"& sup &"' and ANIO_INF='"& anio &"' "&filt&" order by 1,2 desc"
		
	Case "tipo_SupAnioPerSec"
		if sec="1" then
			SQL="Select distinct cod_tipo as cod,descripcion as des from dbo.SMVMA_TIPO t inner join dbo.SMVMA_DIRECTORIO b on t.cod_tipo=b.COD2_EMP and b.anio_inf='"& anio &"' and b.PERINF_EMP='"& per &"' and b.supervision='"& sup &"' and cod_tipo!='00'"
		else sec="2"
			SQL="Select distinct cod_tipo as cod,descripcion as des from dbo.SMVMA_TIPO t where cod_tipo='00'"
		end if
	Case "metodo_SupAnioPerTip"	
		SQL = "SELECT distinct metodo as cod,metodo as des FROM SMVMA_DIRECTORIO where Supervision='"& sup &"' and anio_inf='"& anio &"' and perinf_emp='"& per &"' and cod2_emp='"& tip &"' and moneda_emp!='' order by 1,2 desc"
	Case "moneda_SupAnioPerTip"
		SQL = "SELECT distinct moneda_emp as cod,moneda_emp as des FROM SMVMA_DIRECTORIO where Supervision='"& sup &"' and anio_inf='"& anio &"' and perinf_emp='"& per &"' and cod2_emp='"& tip &"' and moneda_emp!='' order by 1,2 desc"
	Case "moneda_SupAnioPerTipMet"
		SQL = "SELECT distinct moneda_emp as cod,moneda_emp as des FROM SMVMA_DIRECTORIO where Supervision='"& sup &"' and anio_inf='"& anio &"' and perinf_emp='"& per &"' and cod2_emp='"& tip &"' and moneda_emp!='' and metodo='"& met &"' order by 1,2 desc"
	Case "letra_SupAnioPer"
		SQL = "SELECT distinct substring(b.CODCTA_EEFF ,2,1) as cod,case substring(b.CODCTA_EEFF ,2,1) when 'A' then 'A - AFP' when 'F' then 'F - Bancos' when 'E' then 'E - Seguros' when 'C' then 'C - Cavali' when 'V' then 'V - Cavali' when 'D' then 'D - Diversas'  when 'I' then 'I - Agente Bolsa' else substring(b.CODCTA_EEFF ,2,1) end as des FROM "&tabla&" b inner join SMVMA_DIRECTORIO d on isnull(d.si,'0')!='0' and isnull(d.ae14,'0')!='0' and b.ANIO_INF=d.ANIO_INF and b.ID_EMP=d.ID_EMP and b.PERINF_EMP=d.PERINF_EMP and d.Supervision='"&sup&"' and b.TIP_INF=d.TIP_INF where  d.ANIO_INF='"& anio &"' and d.PERINF_EMP='"& per &"' order by 1"
	Case "codigo_SupAnioPerLetrNiv"
		if niv=5 then
			SQL="SELECT distinct d.ae as cod,d.ae +' - ' +d.DescripcionAE as des"
		elseif niv=6 then
			SQL="SELECT distinct d.si as cod,d.si + ' - '+d.descripcionSI as des"
		elseif niv=10 then
			SQL="SELECT distinct d.ae54 as cod,d.ae54 +' - ' +d.DescripcionAE54 as des"
		elseif niv=11 then
			SQL="SELECT distinct d.ae14 as cod,d.ae14 +' - ' +d.DescripcionAE14 as des"
		elseif niv=12 then
			SQL="SELECT distinct COD2_EMP as cod,case when COD2_EMP ='00' then 'NO FINANCIERA' else DESC2_EMP end as des"
		end if

		'Se incluye la tabla SMVMA_ESTGP en esta consulta, ya que no todos los registros
		'que figuran en el directorio tienen estados financieros.
		filtroAE14="and isnull(d.ae14,'0')!='0' "
		if niv=12 then
			filtroAE14=""
		end if

		SQL = SQL +" FROM "&tabla&" b inner join SMVMA_DIRECTORIO d on isnull(d.si,'0')!='0' "+filtroAE14+" and b.ANIO_INF=d.ANIO_INF and b.ID_EMP=d.ID_EMP and b.PERINF_EMP=d.PERINF_EMP and d.Supervision='"&sup&"' and b.TIP_INF=d.TIP_INF where d.ANIO_INF='"& anio &"' and d.PERINF_EMP='"& per &"' and substring(b.CODCTA_EEFF ,2,1)='"& letr &"' order by 1"

		Case "FONDOS_anio"
			SQL = "SELECT distinct ANIO_INF as cod, ANIO_INF as des FROM SMVMA_DIRECTORIO_FONDOS order by 1 desc"

		Case "FONDOS_periodo_Anio"
			SQL = "SELECT distinct PERINF_EMP as cod, case PERINF_EMP when 'A' then 'ANUAL' else 'TRIM '+PERINF_EMP end as des FROM SMVMA_DIRECTORIO_FONDOS where ANIO_INF='"& anio &"' order by 1,2 desc"
	End Select

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQL , con

	Response.ContentType = "application/json; charset=ISO-8859-1"
	
	Response.Write("[")

	If Not rs.BOF Then rs.MoveFirst
	Do While Not rs.EOF

	Response.Write("{ ""cod"": """&rs("cod")&""", ""des"": """&rs("des")&"""},")
	rs.MoveNext
	Loop

	Response.Write("{ }")
	Response.Write("]")

	rs.Close
	Set rs = Nothing
	SQL=""
%>
