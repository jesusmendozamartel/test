function ActualizarMP(){
  	if (document.fmrVarEcon.cboUser.value!="")		
	   {
   			user = document.fmrVarEcon.cboUser.value;
			eretes=document.fmrVarEcon.cboEretes.value;
			ciiu = document.fmrVarEcon.cboCIIU.value;
	   }
	else
		{
			if (document.fmrVarEcon.cboEretes.value!="")
				{	eretes=document.fmrVarEcon.cboEretes.value;
					ciiu = document.fmrVarEcon.cboCIIU.value;
				}
			else
				{	ciiu = document.fmrVarEcon.cboCIIU.value;
				}
		}
 if(document.fmrVarEcon.cboUser.value=="")
	{	user="";
		eretes="";
		ciiu="";
	}
	else
	{
		if (document.fmrVarEcon.cboEretes.value=="")
			{	eretes="";
				ciiu="";
			}
	}
	document.fmrVarEcon.hidUsuario.value=user;
	document.fmrVarEcon.hideretes.value=eretes;
	document.fmrVarEcon.hidCIIU.value=ciiu;
	document.fmrVarEcon.submit();
}

function ActualizarMPSI(){
    tipo = document.fmrSI.cboTipo.value;
	
	document.fmrSI.hidTipo.value=tipo;
	document.fmrSI.submit();
}



//ESTADOR FINACIEROS
function ActualizarEF(){
  	if (document.fmrEF.cboTipo.value!="")		
	   {
   			tipo = document.fmrEF.cboTipo.value;
	/*		sector = document.fmrEF.cboSector.value;
	   }
	else
		{
			sector = document.fmrEF.cboSector.value;
			*/
		
		}
 if(document.fmrEF.cboTipo.value=="")
	{	tipo="";
	//	sector="";
	}
	document.fmrEF.hidTipo.value=tipo;
	//document.fmrEF.hidSector.value=sector;
	document.fmrEF.submit();
}

function ActualizarMe(){
  	if (document.fmrEF.cboMetodo.value!="")		
	   {
   			metodo = document.fmrEF.cboMetodo.value;
	/*		sector = document.fmrEF.cboSector.value;
	   }
	else
		{
			sector = document.fmrEF.cboSector.value;
			*/
		
		}
 if(document.fmrEF.cboMetodo.value=="")
	{	metodo="";
	//	sector="";
	}
	document.fmrEF.hidMetodo.value=metodo;
	//document.fmrEF.hidSector.value=sector;
	document.fmrEF.submit();
}


///////////////////////


function Ajax() 
{
  var xmlHttp=null;
  if (window.ActiveXObject) 
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  else 
    if (window.XMLHttpRequest) 
      xmlHttp = new XMLHttpRequest();
  return xmlHttp;
}

var conexion2;	
var divname;	


function cargaVariableSI(valor){
	
	//alert(valor);
	if (valor==2)
	{
	var xRUC=document.getElementById("RUC").value;
	var Aj1=document.getElementById("Ajuste1").checked;
	var Aj2=document.getElementById("Ajuste2").checked;
	var xformato="";
	var xtipoemp="";
	var xsector="";
	}
	else
	{
		
	var xformato=document.getElementById("cboTipo").value;
	//var xtipoemp=document.getElementById("cboTipEmp").value;
	var xsector=document.getElementById("cboSector").value;
	var Aj1=document.getElementById("Ajuste1").checked;
	var Aj2=document.getElementById("Ajuste2").checked;
	var xRUC="";
	var xRAZ="";
	
	}
	if (valor==2)
	{	if (xRUC=="" )  
		{alert ("Ingrese RUC");}
		else
		{
				//var tam1=document.getElementById("Tam1").checked;
				//var tam2=document.getElementById("Tam2").checked;
				var emes1=document.getElementById("EmEs1").checked;
				var emes2=document.getElementById("EmEs2").checked;
				//var aj1=document.getElementById("Aju1").checked;
				//var aj2=document.getElementById("Aju2").checked;
				
				/*if (tam1==false && tam2==false)
					{alert("Seleccione Empresa Grande o Pequeña");}
				else
				{*/
					if (emes1==false && emes2==false)
						{ alert("Seleccione Empresa O Establecimiento"); }
					else
					{
						if (Aj1==false && Aj2==false)
							{alert("Seleccione Con Ajuste o Sin Ajuste");}
						else
						{
								conexion2=Ajax();
								muestra();
								var url= "SistemaIntermedio_Reporte.asp";
								url=url+"?formato="+xformato+"&tipoemp="+xtipoemp+"&sector="+xsector+"&RUC="+xRUC+"&emes1="+emes1+"&emes2="+emes2+"&aj1="+Aj1+"&aj2="+Aj2;
						alert(url);
								conexion2.open("GET", url, true);
								conexion2.setRequestHeader("Content-Type", "text/html");
								conexion2.setRequestHeader("encoding", "iso-8859-1");
								conexion2.onreadystatechange = procesaVariables;
								conexion2.send(null);
						}
					//}
				
			}
		}
	}
	else
	{
		if (xformato=="" )
		{alert ("Seleccione Formato");}
		else
		{
			if (xtipoemp=="")
			{alert ("Seleccione Empresa o Establecimiento");}
			else
			{
				if	(xsector=="")
					{alert ("Seleccione Sector");}
				else
					{
						//var aj1=document.getElementById("Ajuste1").checked;
						//var aj2=document.getElementById("Ajuste2").checked;
						if (Aj1==false && Aj2==false)
						{alert ("Seleccione Con Ajuste o sin Ajuste");}
						else
						{
						conexion2=Ajax();
						muestra();
						var url= "SistemaIntermedio_Reporte.asp";
						url=url+"?formato="+xformato+"&tipoemp="+xtipoemp+"&sector="+xsector+"&RUC="+xRUC+"&aj1="+Aj1+"&aj2="+Aj2;
						alert(url);
						conexion2.open("GET", url, true);
						conexion2.setRequestHeader("Content-Type", "text/html");
						conexion2.setRequestHeader("encoding", "iso-8859-1");
						conexion2.onreadystatechange = procesaVariables;
						conexion2.send(null);
						}
					}
			}
		}
	}
}

function cargaVariableEF(estadofinanciero, supervision)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=document.getElementById('cboLetra').value.substring(0, 1);
	var xdet=document.getElementById('cboDetalle').value;
	var xmetodo=''
	if (estadofinanciero=='FE'){
		xmetodo=document.getElementById('cboMetodo').value;
	}

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	if(estadofinanciero=='BG' && supervision=='S')
		{var url= 'BalanceGeneral_Reporte.asp';}
	if (estadofinanciero=='EG' && supervision=='S')
		{var url= 'EstGananyPerdidas_Reporte.asp';}
	if (estadofinanciero=='FE' && supervision=='S')
		{var url= 'EstFlujoEfectivo_Reporte.asp';}

	if(estadofinanciero=='BG' && supervision=='NS')
		{var url= 'BalanceGeneral_ReporteNS.asp';}
	if (estadofinanciero=='EG' && supervision=='NS')
		{var url= 'EstGananyPerdidas_ReporteNS.asp';}
	if (estadofinanciero=='FE' && supervision=='NS')
		{var url= 'EstFlujoEfectivo_ReporteNS.asp';}

	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra+'&metodo='+xmetodo+'&detalle='+xdet;;

	conexion2.open('POST', noCache(url), true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}

function cargaVariableEF_FONDOS(estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xmoneda=document.getElementById('cboMoneda').value;
	var xdet=document.getElementById('cboDetalle').value;
	var xTipFondo=document.getElementById('cboFondo').value;
	var xmetodo=''
	if (estadofinanciero=='FE' && xTipFondo=="INV"){
		xmetodo=document.getElementById('cboMetodo').value;
	}

	if (xannio=='')  {
		alert ('Seleccione un Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione un Periodo');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	if(estadofinanciero=='BG')
		{var url= 'BalanceGeneralFONDOS_Reporte.asp';}
	if (estadofinanciero=='GP')
		{var url= 'EstGananyPerdidasFONDOS_Reporte.asp';}
	if (estadofinanciero=='FE')
		{var url= 'EstFlujoEfectivoFONDOS_Reporte.asp';}

	url=url+'?annio='+xannio+'&trime='+xtrim+'&moneda='+xmoneda+'&metodo='+xmetodo+'&detalle='+xdet+'&TipFondo='+xTipFondo;

	conexion2.open('POST', noCache(url), true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}

function cargaVariable(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xorden=document.getElementById('cboOrden').value;
	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}

		if (xorden=='')  {
			alert ('Seleccione el Orden');
			document.getElementById('cboOrden').focus();
			return false;
		}	

		conexion2=Ajax();
		muestra();
		var url= 'Directorio_Reporte.asp';
		
		url=url+'?annio='+xannio+'&orden='+xorden+'&trime='+xtrim;
		//alert(url);
		conexion2.open('POST', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
}

function cargaVariableFONDOS(){
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xTipFondo=document.getElementById('cboFondo').value;
	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}

		if (xTipFondo=='')  {
			alert ('Seleccione el Tipo de Fondo');
			document.getElementById('cboFondo').focus();
			return false;
		}	

		conexion2=Ajax();
		muestra();
		var url= 'DirectorioFondos_Reporte.asp';
		
		url=url+'?annio='+xannio+'&tfondo='+xTipFondo+'&trime='+xtrim;
		//alert(url);
		conexion2.open('POST', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
}


function cargaVariableNS(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xorden=document.getElementById('cboOrden').value;
	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}

		if (xorden=='')  {
			alert ('Seleccione el Orden');
			document.getElementById('cboOrden').focus();
			return false;
		}	

		conexion2=Ajax();
		muestra();
		var url= 'Directorio_ReporteNS.asp';
		
		url=url+'?annio='+xannio+'&orden='+xorden+'&trime='+xtrim;
		//alert(url);
		conexion2.open('POST', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
}


function cargaVariablePN(supervision)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=document.getElementById('cboLetra').value.substring(0, 1);
	
	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	if(supervision=='S')
		{var url= 'CambioPatNeto_Reporte.asp';}
	if (supervision=='NS')
		{var url= 'CambioPatNeto_ReporteNS.asp';}

	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra;

	conexion2.open('POST', noCache(url), true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}

function cargaDetCamPat(CodiEnt)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	
	conexion2=Ajax();
	muestra();
	
	var url= 'CambioPatNeto_Detalle.asp';
	url=url+'?annio='+xannio+'&CodiEnt='+CodiEnt+'&trime='+xtrim

	//alert(url);
	conexion2.open('POST',url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}

function cargaDetCamPatExcel(CodiEnt)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	
	var url= 'CambioPatNeto_DetalleExcel.asp';
	url=url+'?annio='+xannio+'&CodiEnt='+CodiEnt+'&trime='+xtrim

	document.fmrEF.action=url;
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}

function cargaVariablePNNS(valor,estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xsector=document.getElementById('cboSector').value;
	var xtipo=document.getElementById('cboTipo').value;
	
	
	if (valor==1)
	{	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione el Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}

		if (xsector=='')  {
			alert ('Seleccione Sector');			
			document.getElementById('cboSector').focus();
			return false;
		}
		if (xtipo=='')  {
			alert ('Seleccione Tipo');
			document.getElementById('cboTipo').focus();
			return false;
		}

		conexion2=Ajax();
		muestra();
		var url= 'CambioPatNeto_ReporteNS.asp';
				
		url=url+'?annio='+xannio+'&sector='+xsector+'&tipo='+xtipo+'&trime='+xtrim;
		//alert(url);
		conexion2.open('POST', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
	}		
}

function cargaDividendos(supervision)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=document.getElementById('cboLetra').value.substring(0, 1);
	
	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	if(supervision=='S')
		{var url= 'Dividendos_Reporte.asp';}
	if (supervision=='NS')
		{var url= 'Dividendos_ReporteNS.asp';}

	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra;

	conexion2.open('POST', noCache(url), true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}

function cargaConsolidadoEEFF(supervision)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=document.getElementById('cboLetra').value.substring(0, 1);
	var xdet=document.getElementById('cboDetalle').value;
	var xvis=document.getElementById('cboVista').value;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}	
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	if (xvis ==1 && supervision=='S')
		url='ConsolidadoEEFF_Reporte.asp';
	if (xvis ==2 && supervision=='S')
		url='ConsolidadoEEFF2_Reporte.asp';

	if (xvis ==1 && supervision=='NS')
		url='ConsolidadoEEFF_ReporteNS.asp';
	if (xvis ==2 && supervision=='NS')
		url='ConsolidadoEEFF2_ReporteNS.asp';

	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra+'&detalle='+xdet;

	conexion2.open('POST', noCache(url), true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);


	//$.ajax({ url:noCache(url), cache : false, success: function(data){
	    //document.getElementById('DivVariables').innerHTML = data.foo;
	//}, dataType: "json"});	
}

function ExcelConsolidadoEEFF(supervision)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	//var xcod=document.getElementById('cboCodigo').value;
	var xvis=document.getElementById('cboVista').value;

	var cboCod=document.getElementById('cboCodigo');
	var cboLet=document.getElementById('cboLetra');
	var cboDet=document.getElementById('cboDetalle');

	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=cboLet.value.substring(0, 1);
	var xdet=cboDet.value;	

	var xcodText= cboCod.options[cboCod.selectedIndex].text;
	var xletText= cboLet.options[cboLet.selectedIndex].text;
	var xdetText= cboDet.options[cboDet.selectedIndex].text;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}	
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();

	var url= ''

	if (xvis ==1 && supervision=='S')
		url='ConsolidadoEEFF_ReporteExcel.asp';
	if (xvis ==2 && supervision=='S')
		url='ConsolidadoEEFF2_ReporteExcel.asp';

	if (xvis ==1 && supervision=='NS')
		url='ConsolidadoEEFF_ReporteExcelNS.asp';
	if (xvis ==2 && supervision=='NS')
		url='ConsolidadoEEFF2_ReporteExcelNS.asp';

	//var url= 'ConsolidadoEEFF_ReporteExcel.asp';
	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra+'&detalle='+xdet+'&codText='+xcodText+'&detText='+xdetText+'&letText='+xletText;

	document.fmrEF.action=noCache(url);
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}


function ExcelDividendos(supervision)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=document.getElementById('cboLetra').value.substring(0, 1);
	
	var cboCod=document.getElementById('cboCodigo');
	var cboLet=document.getElementById('cboLetra');
	
	var xcodText= cboCod.options[cboCod.selectedIndex].text;
	var xletText= cboLet.options[cboLet.selectedIndex].text;


	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	var url= ''

	if(supervision=='S')
		{var url= 'Dividendos_ReporteExcel.asp';}
	if (supervision=='NS')
		{var url= 'Dividendos_ReporteExcelNS.asp';}

	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra+'&codText='+xcodText+'&letText='+xletText;

	document.fmrEF.action=noCache(url);
	document.fmrEF.submit();
	document.fmrEF.target='_self';		
}

function cargaConsolidadoEEFF_FONDOS()
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xmoneda=document.getElementById('cboMoneda').value;
	var xdet=document.getElementById('cboDetalle').value;
	var xTipFondo=document.getElementById('cboFondo').value;
	var xvis=document.getElementById('cboVista').value;

	if (xannio=='')  {
		alert ('Seleccione un Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione un Periodo');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	if (xvis ==1)
		url='ConsolidadoEEFFFondos_Reporte.asp';
	if (xvis ==2)
		url='ConsolidadoEEFFFondos2_Reporte.asp';

	url=url+'?annio='+xannio+'&trime='+xtrim+'&moneda='+xmoneda+'&detalle='+xdet+'&TipFondo='+xTipFondo;

	conexion2.open('POST', noCache(url), true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}


function ExcelConsolidadoEEFF_FONDOS()
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var cboDet=document.getElementById('cboDetalle');
	var xmoneda=document.getElementById('cboMoneda').value;
	var xdet=document.getElementById('cboDetalle').value;
	var cboFondo=document.getElementById('cboFondo');
	var xTipFondo=document.getElementById('cboFondo').value;
	var xdetText= cboDet.options[cboDet.selectedIndex].text;
	var xdetFondo= cboFondo.options[cboFondo.selectedIndex].text;
	var xvis=document.getElementById('cboVista').value;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	var url= ''

	if (xvis ==1)
		url='ConsolidadoEEFFFondos_ReporteExcel.asp';
	if (xvis ==2)
		url='ConsolidadoEEFFFondos2_ReporteExcel.asp';

	url=url+'?annio='+xannio+'&trime='+xtrim+'&moneda='+xmoneda+'&detalle='+xdet+'&TipFondo='+xTipFondo+'&detText='+xdetText+'&xdetFondo='+xdetFondo;
	//alert(url);
	document.fmrEF.action=noCache(url);
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}


function cargaConsolidadoEEFF2()
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=document.getElementById('cboLetra').value.substring(0, 1);
	var xdet=document.getElementById('cboDetalle').value;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}	
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= 'ConsolidadoEEFF2_Reporte.asp';

	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra+'&detalle='+xdet;

	conexion2.open('POST', noCache(url), true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);


	//$.ajax({ url:noCache(url), cache : false, success: function(data){
	    //document.getElementById('DivVariables').innerHTML = data.foo;
	//}, dataType: "json"});	
}

function ExcelConsolidadoEEFF2()
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	//var xcod=document.getElementById('cboCodigo').value;

	var cboCod=document.getElementById('cboCodigo');
	var cboLet=document.getElementById('cboLetra');
	var cboDet=document.getElementById('cboDetalle');

	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=cboLet.value.substring(0, 1);
	var xdet=cboDet.value;	

	var xcodText= cboCod.options[cboCod.selectedIndex].text;
	var xletText= cboLet.options[cboLet.selectedIndex].text;
	var xdetText= cboDet.options[cboDet.selectedIndex].text;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	conexion2=Ajax();

	var url= 'ConsolidadoEEFF2_ReporteExcel.asp';
	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra+'&detalle='+xdet+'&codText='+xcodText+'&detText='+xdetText+'&letText='+xletText;

	document.fmrEF.action=noCache(url);
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}

function cargaDirVariable(){
	
	var strAnio="";
	var strMoneda="";
	var strNivel="";
	
	strAnio = document.getElementById("cboAnio").value;
	strMoneda = document.getElementById("cboMoneda").value;
	strNivel = document.getElementById("cboNivel").value;

	if (strAnio=='') {alert ('Seleccione Año');return false;}
	if (strMoneda=='') {alert ('Seleccione Moneda');return false;}

	conexion2=Ajax();
	muestra();
	var url= 'DirectorioVariables_Reporte.asp'+"?anio="+strAnio+"&moneda="+strMoneda+"&nivel="+strNivel;
	conexion2.open('POST', url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;

	conexion2.send(null);
}

function cargaExcelDirVariable(){
	
	var strAnio="";
	var strMoneda="";
	var strNivel="";

	var NivelText="";
	
	strAnio = document.getElementById("cboAnio").value;
	strMoneda = document.getElementById("cboMoneda").value;
	strNivel = document.getElementById("cboNivel").value;

	NivelText = $('#cboNivel option:selected').text();

	if (strAnio=='') {alert ('Seleccione Año');return false;}
	if (strMoneda=='') {alert ('Seleccione Moneda');return false;}

	conexion2=Ajax();
	document.fmrEF.action='DirectorioVariables_ReporteExcel.asp'+"?anio="+strAnio+"&moneda="+strMoneda+"&nivel="+strNivel+"&NivelText="+NivelText;
	//document.fmrEF.action='MultipleExcelSheet.asp';
	
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}

function  procesaVariables(){
 //alert("entro");
   if (conexion2.readyState == 4){
	   if(conexion2.responseText == 1 || conexion2.responseText == 0 ){
		  	alert("No se encontraron registros!");
		}else{
   //alert(conexion2.responseText);
		   	document.getElementById('DivVariables').innerHTML = conexion2.responseText;			
		}
	oculta();
  } 
}

function noCache(uri){return uri.concat(/\?/.test(uri)?"&":"?","noCache=",(new Date).getTime(),".",Math.random()*1234567)};

function muestra(){
	//alert("muestra");
	document.getElementById('blocker').style.display="block";
}

function oculta(){
	//	alert("oculta");
	document.getElementById('blocker').style.display="none";
}
function Activa(valor){
	if (valor=="M")
	{
	 document.getElementById("EmEs2").disabled= true;
	 document.getElementById("Aju2").disabled= true;
	}
	else
	{
	 document.getElementById("EmEs2").disabled= false;
	 document.getElementById("Aju2").disabled= false;
		}
	
}
	

function Procesar()
{
var xannio=document.getElementById("cboAnio").value;
var xtipo=document.getElementById("hidTipo").value;
var url="Procesar_Reporte.asp"+"?annio="+xannio+"&tipo="+xtipo;

conexion2=Ajax();
muestra();
conexion2.open("GET", url, true);
conexion2.setRequestHeader("Content-Type", "text/html");
conexion2.setRequestHeader("encoding", "iso-8859-1");
conexion2.onreadystatechange = procesaVariables;

conexion2.send(null);
}


function ExcelEF(estadofinanciero, supervision)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var cboCod=document.getElementById('cboCodigo');
	var cboLet=document.getElementById('cboLetra');
	var cboDet=document.getElementById('cboDetalle');
	var xmoneda=document.getElementById('cboMoneda').value;
	var xletra=document.getElementById('cboLetra').value.substring(0, 1);
	var xdet=document.getElementById('cboDetalle').value;
	var xmetodo=''
	var xcodText= cboCod.options[cboCod.selectedIndex].text;
	var xletText= cboLet.options[cboLet.selectedIndex].text;
	var xdetText= cboDet.options[cboDet.selectedIndex].text;

	if (estadofinanciero=='FE'){
		xmetodo=document.getElementById('cboMetodo').value;
	}

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	var url= ''
	if(estadofinanciero=='BG' && supervision=='S')
		{var url= 'BalanceGeneral_ReporteExcel.asp';}
	if (estadofinanciero=='EG' && supervision=='S')
		{var url= 'EstGananyPerdidas_ReporteExcel.asp';}
	if (estadofinanciero=='FE' && supervision=='S')
		{var url= 'EstFlujoEfectivo_ReporteExcel.asp';}

	if(estadofinanciero=='BG' && supervision=='NS')
		{var url= 'BalanceGeneral_ReporteExcelNS.asp';}
	if (estadofinanciero=='EG' && supervision=='NS')
		{var url= 'EstGananyPerdidas_ReporteExcelNS.asp';}
	if (estadofinanciero=='FE' && supervision=='NS')
		{var url= 'EstFlujoEfectivo_ReporteExcelNS.asp';}

	url=url+'?annio='+xannio+'&trime='+xtrim+'&nivel='+xNiv+'&codigo='+xcod+'&moneda='+xmoneda+'&letra='+xletra+'&metodo='+xmetodo+'&detalle='+xdet+'&codText='+xcodText+'&detText='+xdetText+'&letText='+xletText;
	//alert(url);
	document.fmrEF.action=noCache(url);
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}


function ExcelEF_FONDOS(estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var cboDet=document.getElementById('cboDetalle');
	var xmoneda=document.getElementById('cboMoneda').value;
	var xdet=document.getElementById('cboDetalle').value;
	var xTipFondo=document.getElementById('cboFondo').value;
	var xmetodo=''
	var xdetText= cboDet.options[cboDet.selectedIndex].text;

	if (estadofinanciero=='FE'){
		xmetodo=document.getElementById('cboMetodo').value;
	}

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	if (xtrim=='')  {
		alert ('Seleccione Trimestre');
		document.getElementById('cboTrim').focus();
		return false;
	}
	if (xmoneda=='')  {
		alert ('Seleccione la Moneda');
		document.getElementById('cboMoneda').focus();
		return false;
	}

	var url= ''
	if(estadofinanciero=='BG')
		{var url= 'BalanceGeneralFondos_ReporteExcel.asp';}
	if (estadofinanciero=='GP')
		{var url= 'EstGananyPerdidasFondos_ReporteExcel.asp';}
	if (estadofinanciero=='FE')
		{var url= 'EstFlujoEfectivoFondos_ReporteExcel.asp';}

	url=url+'?annio='+xannio+'&trime='+xtrim+'&moneda='+xmoneda+'&metodo='+xmetodo+'&detalle='+xdet+'&TipFondo='+xTipFondo+'&detText='+xdetText;
	//alert(url);
	document.fmrEF.action=noCache(url);
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}

function ExcelEFNS(valor,estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xsector=document.getElementById('cboSector').value;
	var xtipo=document.getElementById('cboTipo').value;
	var xorden=document.getElementById('cboOrden').value;	
	var xmoneda=document.getElementById('cboMoneda').value;

	var xmetodo='';

	if(estadofinanciero!='BG' && estadofinanciero!='EG') {
		xmetodo=document.getElementById('cboMetodo').value;
	}

	if (valor==1)	{	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xsector=='')  {
			alert ('Seleccione Área');			
			document.getElementById('cboSector').focus();
			return false;
		}
		if (xtipo=='')  {
			alert ('Seleccione Moneda');
			document.getElementById('cboTipo').focus();
			return false;
		}
		if (xorden=='')  {
			alert ('Seleccione uno o mas Pliegos');
			document.getElementById('cboOrden').focus();
			return false;
		}

		if (estadofinanciero=='BG') {
		document.fmrEF.action='BalanceGeneral_ReporteExcelNS.asp'+"?annio="+xannio+"&sector="+xsector+"&tipo="+xtipo+"&orden="+xorden+'&trime='+xtrim+'&metodo='+xmetodo+'&moneda='+xmoneda;
		document.fmrEF.submit();
		//document.fmrEF.action="BalanceGeneral.asp";
		//document.fmrEF.target='_self';
		}
		
		if (estadofinanciero=='EG') { 
		document.fmrEF.action='EstGananyPerdidas_ReporteExcelNS.asp'+"?annio="+xannio+"&sector="+xsector+"&tipo="+xtipo+"&orden="+xorden+'&trime='+xtrim+'&metodo='+xmetodo+'&moneda='+xmoneda;
		document.fmrEF.submit();
		//document.fmrEF.action="EstGananyPerdidas.asp";
		//document.fmrEF.target='_self';
		}
		if (estadofinanciero=='FE') {
		document.fmrEF.action='EstFlujoEfectivo_ReporteExcelNS.asp'+"?annio="+xannio+"&sector="+xsector+"&tipo="+xtipo+"&orden="+xorden+'&trime='+xtrim+'&metodo='+xmetodo+'&moneda='+xmoneda;
		document.fmrEF.submit();
		//document.fmrEF.action="EstFlujoEfectivo.asp";
		//document.fmrEF.target='_self';
		}
	}
}

function ExcelPN(valor,estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xsector=document.getElementById('cboSector').value;
	var xtipo=document.getElementById('cboTipo').value;
	
	if (valor==1)	{	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xsector=='')  {
			alert ('Seleccione Área');			
			document.getElementById('cboSector').focus();
			return false;
		}
		if (xtipo=='')  {
			alert ('Seleccione Moneda');
			document.getElementById('cboTipo').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione el Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}		
		
		if (estadofinanciero=='PN') { 
		document.fmrEF.action='CambioPatNeto_ReporteExcel.asp'+"?annio="+xannio+"&sector="+xsector+"&tipo="+xtipo+"&trime="+xtrim;
		document.fmrEF.submit();
		document.fmrEF.action="CambioPatNeto.asp";
		document.fmrEF.target='_self';
		}

	}
}
function ExcelPNNS(valor,estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xtrim=document.getElementById('cboTrim').value;
	var xsector=document.getElementById('cboSector').value;
	var xtipo=document.getElementById('cboTipo').value;
	
	if (valor==1)	{	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xsector=='')  {
			alert ('Seleccione Área');			
			document.getElementById('cboSector').focus();
			return false;
		}
		if (xtipo=='')  {
			alert ('Seleccione Moneda');
			document.getElementById('cboTipo').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione el Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}		
		
		if (estadofinanciero=='PN') { 
		document.fmrEF.action='CambioPatNeto_ReporteExcelNS.asp'+"?annio="+xannio+"&sector="+xsector+"&tipo="+xtipo+"&trime="+xtrim;
		document.fmrEF.submit();
		document.fmrEF.action="CambioPatNetoNS.asp";
		document.fmrEF.target='_self';
		}

	}
}
function Excel(valor,estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xorden=document.getElementById('cboOrden').value;
	var xtrim=document.getElementById('cboTrim').value;
	
	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}

		if (xorden=='')  {
			alert ('Seleccione el Orden');
			document.getElementById('cboOrden').focus();
			return false;
		}		
		
		document.fmrEF.action='Directorio_ReporteExcel.asp'+"?annio="+xannio+"&orden="+xorden+"&trime="+xtrim;
		document.fmrEF.submit();
		document.fmrEF.action="Directorio.asp";
		document.fmrEF.target='_self';
	
}
function ExcelFONDOS()
{
	var xannio=document.getElementById('cboAnio').value;
	var xtipfondos=document.getElementById('cboFondo').value;
	var xtrim=document.getElementById('cboTrim').value;
	
	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}

		if (xtipfondos=='')  {
			alert ('Seleccione el Orden');
			document.getElementById('cboFondo').focus();
			return false;
		}		
		
		document.fmrEF.action='DirectorioFondos_ReporteExcel.asp'+"?annio="+xannio+"&xtipfondos="+xtipfondos+"&trime="+xtrim;
		document.fmrEF.submit();
		document.fmrEF.action="DirectorioFondos.asp";
		document.fmrEF.target='_self';
	
}

function ExcelNS(valor,estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xorden=document.getElementById('cboOrden').value;
	var xtrim=document.getElementById('cboTrim').value;
	
	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xtrim=='')  {
			alert ('Seleccione Trimestre');
			document.getElementById('cboTrim').focus();
			return false;
		}

		if (xorden=='')  {
			alert ('Seleccione el Orden');
			document.getElementById('cboOrden').focus();
			return false;
		}		
		
		document.fmrEF.action='Directorio_ReporteExcelNS.asp'+"?annio="+xannio+"&orden="+xorden+"&trime="+xtrim;
		document.fmrEF.submit();
		document.fmrEF.action="DirectorioNS.asp";
		document.fmrEF.target='_self';
	
}
function alerta()
{
	alert ("No hay ningun registro");
}

function CargaFiltro(cboName,url,FuncDepend)
{
	firstValue=0
	var c = document.getElementById(cboName);
	$.ajax({
	    url: url,
	    type: 'POST',
	    success: function(data) {
			c.innerHTML = "";
	    	for (var i = 0; i < data.length-1; i++) {
				var option = document.createElement("option");
				option.text = data[i].des;					
				option.value =  data[i].cod;
				c.add(option);

	    	}
			if (FuncDepend!="")
			{
    			call_others(FuncDepend);
			}
	    },
	    error: function() {
	    	alert("Ocurrió un error. Comuníquese con el administrador del sistema.");
	    },
	    cache: false,contentType: "json; charset:ISO-8859-1",processData: false
	}, 'json');
}

function call_others(function_name) {
	window[function_name]();
	//eval(function_name+"()");
}

function ActualizarEF()
{
	if (document.fmrEF.cboTipo.value!="")		
	   {
   			tipo = document.fmrEF.cboTipo.value;
	   }
	else
		{	tipo="";
		}
	document.fmrEF.hidTipo.value=tipo;
	document.fmrEF.submit();
}


function Refresh(valor)
{
	
	if (valor==1) {document.URL = "Directorio.asp";}
	if (valor==2) {document.URL = "BalanceGeneral.asp";}
	if (valor==3) {document.URL = "EstGananyPerdidas.asp";}
	if (valor==4) {document.URL = "EstFlujoEfectivo.asp";}
	if (valor==5) {document.URL = "EstadoCambioPatNeto.asp";}
	if (valor==6) {document.URL = "BalanceGeneralNS.asp";}
	if (valor==7) {document.URL = "DirectorioNS.asp";}
	if (valor==8) {document.URL = "EstGananyPerdidasNS.asp";}
	if (valor==9) {document.URL = "EstFlujoEfectivoNS.asp";}
	if (valor==10) {document.URL = "CambioPatNetoNS.asp";}
	if (valor==11) {document.URL = "DirectorioVariables.asp";}
}


function cargaArchivos(Rep,filt){

	var cbo=document.getElementById('cboAnio');
	var xannio=''

	if (null != cbo) {
    	xannio=cbo.value;
    }
	conexion2=Ajax();
	muestra();

	var url= '';
	var jerarquia= '';
	url= 'ListFiles.asp?anio='+xannio+'&jer='+Rep+'&filt='+filt;

	//PDT ANUAL
	//PDTA_MD->Metadatos
	//if(Rep=="PDTA_MD")	url= 'Metadatos_Anual_Reporte.asp?anio='+xannio+'&jer=02';
	//PDTA_DP->Directorio Paneles
	//if(Rep=="PDTA_DP")	url= 'DirPaneles_Anual_Reporte.asp?anio='+xannio+'&jer=03';
	//if(Rep=="PDTA_DP")	url= 'ListFiles.asp?anio='+xannio+'&jer=03&tipo=1';
	//PDTA_EF->Estados Financieros
	//PDTA_SI->Sistema Intermedio
	//if(Rep=="PDTA_SI")	url= 'ListFiles.asp?anio='+xannio+'&jer=04&tipo=1';
	//PDTA_RV->Reporte de ventas
	//PDTA_OR->Otros Reportes
	//if(Rep=="PDTA_OR")	url= 'OtrosReportesAnual_Reporte.asp?anio='+xannio+'&jer=01';
	//PDTA_DO->Documentacion
	//if(Rep=="PDTA_DO")	url= 'ListFiles.asp?jer=05&tipo=2';

	//PDT MENSUAL
	//PDTM_MD->Metadatos
	//PDTM_DP->Directorio Paneles
	//PDTM_RV->Registro de Ventas y Compras
	//PDTM_RE->Reportes
	//PDTM_EF->Estados Financieros: Libros Electrónicos
	//PDTM_SI->Sistema Intermedio
	//PDTM_DO->Documentación

	//DIRECTORIO
	//DIR_MD->Metadatos
	//DIR_AN->Anual
	//DIR_ME->Mensual
	//DIR_LG->Listado General - Renta de Tercera Categoría

	//url=url+'?anio='+xannio+'&jer=01';
	//alert (url);
	conexion2.open('POST', url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);
}
