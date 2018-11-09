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
  	if (document.fmrSI.cboTipo.value!="")		
	   {
   			tipo = document.fmrSI.cboTipo.value;
			tipoemp=document.fmrSI.cboTipEmp.value;
			sector = document.fmrSI.cboSector.value;
	   }
	else
		{
			if (document.fmrSI.cboTipEmp.value!="")
				{	tipoemp=document.fmrSI.cboTipEmp.value;
					sector = document.fmrSI.cboSector.value;
				}
			else
				{	sector = document.fmrSI.cboSector.value;
				}
		}
 if(document.fmrSI.cboTipo.value=="")
	{	tipo="";
		tipoemp="";
		sector="";
	}
	else
	{
		if (document.fmrSI.cboTipEmp.value=="")
			{	tipoemp="";
				sector="";
			}
	}
	
	//alert (tipo);
	//alert (tipoemp);
	//alert (sector);
	
	//alert (document.fmrSI.hidTipo);
	document.fmrSI.hidTipo.value=tipo;
	document.fmrSI.hidTipEmp.value=tipoemp;
	document.fmrSI.hidSector.value=sector;
	document.fmrSI.submit();
}



//ESTADOR FINACIEROS


function ActualizarBG(){
  	if (document.fmrBG.cboTipo.value!="")		
	   {
   			tipo = document.fmrBG.cboTipo.value;
			//tipoemp=document.fmrBG.cboTipEmp.value;
			sector = document.fmrBG.cboSector.value;
	   }
	else
		{
			sector = document.fmrBG.cboSector.value;
		
		}
 if(document.fmrBG.cboTipo.value=="")
	{	tipo="";
		//tipoemp="";
		sector="";
	}

	//alert (tipo);
	//alert (tipoemp);
	//alert (sector);
	
	//alert (document.fmrSI.hidTipo);
	document.fmrBG.hidTipo.value=tipo;
	//document.fmrSI.hidTipEmp.value=tipoemp;
	document.fmrBG.hidSector.value=sector;
	document.fmrBG.submit();
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
//
function cargaVariable(valor){
	
	//alert(valor);
	if (valor==2)
	{
	var xRUC=document.getElementById("RUC").value;
	var xRAZ=document.getElementById("RSocial").value;
	var xusuario="";
	var xeretes="";
	var xciiu="";
	}
	else
	{
	var xusuario=document.getElementById("cboUser").value;
	var xeretes=document.getElementById("cboEretes").value;
	var xciiu=document.getElementById("cboCIIU").value;
	var xRUC="";
	var xRAZ="";
	}
	if (valor==2)
	{	if (xRUC=="" && xRAZ=="")  
		{alert ("Ingrese RUC o Razon Social");}
		else
		{
			conexion2=Ajax();
			muestra();
			var url= "ReporteVariables.asp";
			url=url+"?usuario="+xusuario+"&eretes="+xeretes+"&ciiu="+xciiu+"&RUC="+xRUC+"&RAZ="+xRAZ;
			//alert(url);
			conexion2.open("GET", url, true);
			conexion2.setRequestHeader("Content-Type", "text/html");
			conexion2.setRequestHeader("encoding", "iso-8859-1");
			conexion2.onreadystatechange = procesaVariables;
			conexion2.send(null);
		}
	}
	else
	{
	conexion2=Ajax();
	muestra();
	var url= "ReporteVariables.asp";
	url=url+"?usuario="+xusuario+"&eretes="+xeretes+"&ciiu="+xciiu+"&RUC="+xRUC+"&RAZ="+xRAZ;
	//alert(url);
	conexion2.open("GET", url, true);
	conexion2.setRequestHeader("Content-Type", "text/html");
	conexion2.setRequestHeader("encoding", "iso-8859-1");
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);
	}
}

function cargaVariableSI(valor){
	
	//alert(valor);
	if (valor==2)
	{
	var xRUC=document.getElementById("RUC").value;
	var xRAZ=document.getElementById("RSocial").value;
	var xformato="";
	var xtipoemp="";
	var xsector="";
	}
	else
	{
	var xformato=document.getElementById("cboTipo").value;
	var xtipoemp=document.getElementById("cboTipEmp").value;
	var xsector=document.getElementById("cboSector").value;
	var xRUC="";
	var xRAZ="";
	}
	if (valor==2)
	{	if (xRUC=="" && xRAZ=="")  
		{alert ("Ingrese RUC o Razon Social");}
		else
		{
			if (xRAZ.length<5)
				{alert ("Escriba mas letras");}
			
			else
			{
				var tam1=document.getElementById("Tam1").checked;
				var tam2=document.getElementById("Tam2").checked;
				var emes1=document.getElementById("EmEs1").checked;
				var emes2=document.getElementById("EmEs2").checked;
				var aj1=document.getElementById("Aju1").checked;
				var aj2=document.getElementById("Aju2").checked;
				
				if (tam1==false && tam2==false)
					{alert("Seleccione Empresa Grande o Pequeña");}
				else
				{
					if (emes1==false && emes2==false)
						{ alert("Seleccione Empresa O Establecimiento"); }
					else
					{
						if (aj1==false && aj2==false)
							{alert("Seleccione Con Ajuste o Sin Ajuste");}
						else
						{
								conexion2=Ajax();
								muestra();
								var url= "ReportSistemaIntermedio.asp";
								url=url+"?formato="+xformato+"&tipoemp="+xtipoemp+"&sector="+xsector+"&RUC="+xRUC+"&RAZ="+xRAZ+"&tam1="+tam1+"&tam2="+tam2+"&emes1="+emes1+"&emes2="+emes2+"&aj1="+aj1+"&aj2="+aj2;
						//alert(url);
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
					var aj1=document.getElementById("Ajuste1").checked;
					var aj2=document.getElementById("Ajuste2").checked;
						if (aj1==false && aj2==false)
						{alert ("Seleccione Con Ajuste o sin Ajuste");}
						else
						{
						conexion2=Ajax();
						muestra();
						var url= "ReportSistemaIntermedio.asp";
						url=url+"?formato="+xformato+"&tipoemp="+xtipoemp+"&sector="+xsector+"&RUC="+xRUC+"&RAZ="+xRAZ;
						//alert(url);
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


/////////////

function cargaVariableBG(valor){
	
	//alert(valor);
	if (valor==2)
	{
	var xRUC=document.getElementById("RUC").value;
	//var xRAZ=document.getElementById("RSocial").value;
	var xformato="";
	var xsector="";
	}
	else
	{
	var xformato=document.getElementById("cboTipo").value;
	var xsector=document.getElementById("cboSector").value;
	var xRUC="";
	//var xRAZ="";
	}
	if (valor==2)
	{	if (xRUC=="")  
		{alert ("Ingrese RUC");}
		else
		{
				var tam1=document.getElementById("Tam1").checked;
				var tam2=document.getElementById("Tam2").checked;
				var aj1=document.getElementById("Aju1").checked;
				var aj2=document.getElementById("Aju2").checked;
								
				if (tam1==false && tam2==false)
					{alert("Seleccione Empresa Grande o Pequeña");}
				else
				{
					if (aj1==false && aj2==false)
						{alert("Seleccione Con Ajuste o Sin Ajuste");}
					else
						{
							conexion2=Ajax();
							muestra();
							var url= "BalanceGeneral_Reporte.asp";
							url=url+"?formato="+xformato+"&sector="+xsector+"&RUC="+xRUC+"&tam1="+tam1+"&tam2="+tam2+"&aj1="+aj1+"&aj2="+aj2;
						//alert(url);
							conexion2.open("GET", url, true);
							conexion2.setRequestHeader("Content-Type", "text/html");
							conexion2.setRequestHeader("encoding", "iso-8859-1");
							conexion2.onreadystatechange = procesaVariables;
							conexion2.send(null);
						}
					
				}
			
		}
	}
	else
	{
		if (xformato=="" )
		{alert ("Seleccione Formato");}
		else
		{
			if	(xsector=="")
			{alert ("Seleccione Sector");}
				else
				{
				var aj1=document.getElementById("Ajuste1").checked;
				var aj2=document.getElementById("Ajuste2").checked;
				if (aj1==false && aj2==false)
					{alert ("Seleccione Con Ajuste o sin Ajuste");}
					else
					{
					conexion2=Ajax();
					muestra();
					var url= "BalanceGeneral_Reporte.asp";
					url=url+"?formato="+xformato+"&sector="+xsector+"&RUC="+xRUC;
					//alert(url);
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



/*
function cargaVariable2(){
	//alert ("PRUEBA");
	var xusuario="";
	var xeretes="";
	var xciiu="";
	var xRUC=document.getElementById("RUC").value;
	var xRAZ=document.getElementById("RSocial").value;
	if (xRUC=="" && xRAZ=="")  
		{alert ("Ingrese RUC o Razon Social");}
	else
	{
		conexion2=Ajax();
		muestra();
		var url= "ReporteVariables.asp";
		url=url+"?usuario="+xusuario+"&eretes="+xeretes+"&ciiu="+xciiu+"&RUC="+xRUC+"&RAZ="+xRAZ;
		//alert(url);
		conexion2.open("GET", url, true);
		conexion2.setRequestHeader("Content-Type", "text/html");
		conexion2.setRequestHeader("encoding", "iso-8859-1");
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
	}
}
*/
function  procesaVariables(){
// alert("entro");
   if (conexion2.readyState == 4){
	   if(conexion2.responseText == 1 || conexion2.responseText == 0 ){
		  	alert("No se encontraron registros!");
		}else{
  // alert(conexion2.responseText);
		   	document.getElementById('DivVariables').innerHTML = conexion2.responseText;
		}
	oculta();
  } 
}


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
	


function Excel(valor)
{
	//alert(valor);
	if (valor==2)
	{
		var xRUC=document.getElementById("RUC").value;
		var xRAZ=document.getElementById("RSocial").value;
		var xusuario="";
		var xeretes="";
		var xciiu="";
	}
	else
	{
		var xusuario=document.getElementById("cboUser").value;
		var xeretes=document.getElementById("cboEretes").value;
		var xciiu=document.getElementById("cboCIIU").value;
		var xRUC="";
		var xRAZ="";
	}
	if (valor==2)
	{
		if (xRUC=="" && xRAZ=="")  
			{alert ("Ingrese RUC o Razon Social");}
		else
			{	document.fmrVarEcon.action='ReporteVariables_Excel.asp';
			//document.fmrVarEcon.target='filtros';
			document.fmrVarEcon.submit();
			document.fmrVarEcon.action="Variables_Economicas.asp";
			document.fmrVarEcon.target='_self';
			
			}
	}
	else
	{
			document.fmrVarEcon.action='ReporteVariables_Excel.asp';
			//document.fmrVarEcon.target='filtros';
			document.fmrVarEcon.submit();
			document.fmrVarEcon.action="Variables_Economicas.asp";
			document.fmrVarEcon.target='_self';
	}
}



function ExcelSI(valor)
{
	//alert(valor);
	if (valor==2)
	{
	var xRUC=document.getElementById("RUC").value;
	var xRAZ=document.getElementById("RSocial").value;
	var tam1=document.getElementById("Tam1").checked;
	var tam2=document.getElementById("Tam2").checked;
	var emes1=document.getElementById("EmEs1").checked;
	var emes2=document.getElementById("EmEs2").checked;
	var aj1=document.getElementById("Aju1").checked;
	var aj2=document.getElementById("Aju2").checked;
	//var xRAZ=document.getElementById("RSocial").value;
	var xformato="";
	var xtipoemp="";
	var xsector="";
	}
	else
	{
	var xformato=document.getElementById("cboTipo").value;
	var xtipoemp=document.getElementById("cboTipEmp").value;
	var xsector=document.getElementById("cboSector").value;
	var xRUC="";
	var xRAZ="";
	}
	if (valor==2)
	{
		if (xRUC=="" && xRAZ=="")  
			{alert ("Ingrese RUC o Razon Social");}
		else
			{
				if (xRUC=="" && (xRAZ.length)<5)
			    	{alert ("Escriba mas letras");}
				else
					{
						if (tam1=="" && tam2=="")
							{alert ("Seleccione tamaño de Empresa");}
						else
							{
								if (emes1=="" && emes2=="")
									{alert ("Seleccione Empresa o establecimiento");}
								else
									{
										if (aj1=="" && aj2=="")
											{alert ("seleccione Con Ajuste o Sin ajuste");}
										else
											{document.fmrSI.action='ReportSistemaIntermedio_Excel.asp';
											//document.fmrVarEcon.target='filtros';
											document.fmrSI.submit();
											document.fmrSI.action="SistemaIntermedio.asp";
											document.fmrSI.target='_self';
											}
									}
							}
			 		}
			
			}
	}
	else
	{
			var aj1=document.getElementById("Ajuste1").checked;
			var aj2=document.getElementById("Ajuste2").checked;
			if (aj1==false && aj2==false)
				{alert ("Seleccione Con Ajuste o sin Ajuste");}
			else
				{
					document.fmrSI.action='ReportSistemaIntermedio_Excel.asp';
					//document.fmrVarEcon.target='filtros';
					document.fmrSI.submit();
					document.fmrSI.action="SistemaIntermedio.asp";
					document.fmrSI.target='_self';
				}
	}
}




function ExcelBG(valor)
{
	//alert(valor);
	if (valor==2)
	{
	var xRUC=document.getElementById("RUC").value;
	//var xRAZ=document.getElementById("RSocial").value;
	var tam1=document.getElementById("Tam1").checked;
	var tam2=document.getElementById("Tam2").checked;
	//var emes1=document.getElementById("EmEs1").checked;
	//var emes2=document.getElementById("EmEs2").checked;
	var aj1=document.getElementById("Aju1").checked;
	var aj2=document.getElementById("Aju2").checked;
	//var xRAZ=document.getElementById("RSocial").value;
	var xformato="";
	//var xtipoemp="";
	var xsector="";
	}
	else
	{
	var xformato=document.getElementById("cboTipo").value;
	//var xtipoemp=document.getElementById("cboTipEmp").value;
	var xsector=document.getElementById("cboSector").value;
	var xRUC="";
	var xRAZ="";
	}
	if (valor==2)
	{
		if (xRUC=="" )  
			{alert ("Ingrese RUC");}
		else
			{
					if (tam1=="" && tam2=="")
							{alert ("Seleccione tamaño de Empresa");}
					else
							{
								if (aj1=="" && aj2=="")
									{alert ("seleccione Con Ajuste o Sin ajuste");}
								else
									{
									document.fmrBG.action='BalanceGeneral_ReporteExcel.asp';
									//document.fmrVarEcon.target='filtros';
									document.fmrBG.submit();
									document.fmrBG.action="BalanceGeneral.asp";
									document.fmrBG.target='_self';
								}
							}
			 				
			}
	}
	else
	{
			var aj1=document.getElementById("Ajuste1").checked;
			var aj2=document.getElementById("Ajuste2").checked;
			if (aj1==false && aj2==false)
				{alert ("Seleccione Con Ajuste o sin Ajuste");}
			else
				{
					document.fmrBG.action='BalanceGeneral_ReporteExcel.asp';
					//document.fmrVarEcon.target='filtros';
					document.fmrBG.submit();
					document.fmrBG.action="BalanceGeneral.asp";
					document.fmrBG.target='_self';
				}
	}
}






function alerta()
{
	alert ("No hay ningun registro");
}
/*
function ActualizarBoton(valor){
document.getElementById('DivVariables').innerHTML ="";
alert (valor)
if (valor=="1")
 {
	document.getElementById('formulario1').style.display="block";
	//document.getElementById('formulario2').style.display="none";
 }
 else
  {
	document.getElementById('formulario1').style.display="none";
	document.getElementById('formulario2').style.display="block";
 }

}
*/

function ActualizarOpcion(valor)
{
	if (valor==1)
	{
		document.fmrVarEcon.submit();
	}
	else
	{
		if (valor==2)
		{document.fmrSI.submit();}
		else
		{document.fmrBG.submit();}
	}

}

function ActualizarEF()
{
	
	if (document.fmrEF.cboTipo.value!="")		
	   {
   			tipo = document.fmrEF.cboTipo.value;
			ficha= document.fmrEF.ocult.value;
	   }
	else
		{	tipo="";
		}
	
	document.fmrEF.hidTipo.value=tipo;
	document.fmrEF.hidEF.value=ficha;
	document.fmrEF.submit();
}


