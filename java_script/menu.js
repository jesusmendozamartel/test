var ruta;
ruta="imagenes/";

stm_bm(["menu45ba",810,"",ruta+"blank.gif",0,"","",0,0,250,0,1000,1,0,0,"","100%",0,0,1,2,"default","hand",""],this);
stm_bp("p0",[0,4,0,0,0,5,20,9,100,"",-2,"",-2,50,2,3,"#999999","transparent",ruta+"bluefireback1.gif",1,1,1,"#000000 #666666 #498bfb"]);
stm_ai("p0i0",[0,"Inicio","","",-1,-1,0,"principal.asp","_self","","",ruta+"folder_home_ov.png",ruta+"folder_home.png",20,16,0,"","",0,0,0,0,1,"#FFFFF7",1,"#B5BED6",1,"",ruta+"bluefireback22.gif",3,0,0,0,"#FFFFF7","#000000","#FFFFFF","#FFFFFF","9pt Verdana","9pt Verdana",0,0],90,0);

stm_aix("p0i1","p0i0",[0,"Reportes Emp. Supervisadas","","",-1,-1,0,"#","_self","","",ruta+"document_graph_ov.png",ruta+"document_graph.png",16,16,0,ruta+"0604arroldw.gif",ruta+"0604arroldw.gif",9,7,0,0,0,"#FFFFF7",1,"#B5BED6",1,"",ruta+"bluefireback2.gif"],150,0);
	stm_bp("p1",[1,4,0,0,5,5,0,0,85,"",-2,"",-2,50,2,3,"#999999","#333333","",0,1,1,"#498bfb"]);
		stm_aix("p1i0","p0i0",[0,"Empresas","","",-1,-1,0,"#","_self","","","","",0,0,0,"","",0,0,0,0,1,"#00CCFF",1,"#498bfb",0,"","",3,0,0,0,"#FFFFF7","#000000","#FFFFFF","#000000","8pt Verdana","8pt Verdana"]);
			stm_bp("p1",[1,4,70,-22,5,5,0,0,85,"",-2,"",-2,50,2,3,"#999999","#333333","",0,1,1,"#498bfb"]);
			stm_aix("p1i0","p0i0",[0,"Metadatos","","",-1,-1,0,"files.asp?rep=02&filt=1&des=Metadatos","_self","","","","",0,0,0,"","",0,0,0,0,1,"#00CCFF",1,"#498bfb",0,"","",3,0,0,0,"#FFFFF7","#000000","#FFFFFF","#000000","8pt Verdana","8pt Verdana"]);
			stm_aix("p1i0","p0i0",[0,"Directorio","","",-1,-1,0,"Directorio.asp","_self","","","","",0,0,0,"","",0,0,0,0,1,"#00CCFF",1,"#498bfb",0,"","",3,0,0,0,"#FFFFF7","#000000","#FFFFFF","#000000","8pt Verdana","8pt Verdana"]);
			stm_aix("p2i1","p1i0",[0,"Balance General","","",-1,-1,0,"BalanceGeneral.asp"]);
			stm_aix("p2i2","p1i0",[0,"Estado Ganancias y Perdidas","","",-1,-1,0,"EstGananyPerdidas.asp"]);
			stm_aix("p2i3","p1i0",[0,"Flujo de Efectivo","","",-1,-1,0,"EstFlujoEfectivo.asp"]);
			stm_aix("p2i4","p1i0",[0,"Estado Cambio Patrimonio Neto","","",-1,-1,0,"CambioPatNeto.asp"]);
			stm_aix("p2i4","p1i0",[0,"Dividendos","","",-1,-1,0,"Dividendos.asp"]);
			stm_aix("p2i5","p1i0",[0,"Directorio de Variables","","",-1,-1,0,"DirectorioVariables.asp"]);
			stm_aix("p2i6","p1i0",[0,"Consolidado EEFF","","",-1,-1,0,"ConsolidadoEEFF.asp"]);
		stm_ep();

		stm_aix("p2i7","p1i0",[0,"Fondos","","",-1,-1,0,"#"]);
		stm_bp("p1",[1,4,70,-22,5,5,0,0,85,"",-2,"",-2,50,2,3,"#999999","#333333","",0,1,1,"#498bfb"]);
			stm_aix("p1i0","p0i0",[0,"Metadatos","","",-1,-1,0,"files.asp?rep=02&filt=1&des=Metadatos","_self","","","","",0,0,0,"","",0,0,0,0,1,"#00CCFF",1,"#498bfb",0,"","",3,0,0,0,"#FFFFF7","#000000","#FFFFFF","#000000","8pt Verdana","8pt Verdana"]);
			stm_aix("p1i0","p0i0",[0,"Directorio","","",-1,-1,0,"DirectorioFONDOS.asp","_self","","","","",0,0,0,"","",0,0,0,0,1,"#00CCFF",1,"#498bfb",0,"","",3,0,0,0,"#FFFFF7","#000000","#FFFFFF","#000000","8pt Verdana","8pt Verdana"]);
			stm_aix("p2i1","p1i0",[0,"Balance General","","",-1,-1,0,"BalanceGeneralFondos.asp"]);
			stm_aix("p2i2","p1i0",[0,"Estado Ganancias y Perdidas","","",-1,-1,0,"EstGananyPerdidasFONDOS.asp"]);
			stm_aix("p2i3","p1i0",[0,"Flujo de Efectivo","","",-1,-1,0,"EstFlujoEfectivoFONDOS.asp"]);
			stm_aix("p2i4","p1i0",[0,"Consolidado EEFF","","",-1,-1,0,"ConsolidadoEEFFFondos.asp"]);
		stm_ep();
stm_ep();
stm_aix("p0i1","p0i0",[0,"Reportes Emp. no Supervisadas","","",-1,-1,0,"#","_self","","",ruta+"document_graph_ov.png",ruta+"document_graph.png",16,16,0,ruta+"0604arroldw.gif",ruta+"0604arroldw.gif",9,7,0,0,0,"#FFFFF7",1,"#B5BED6",1,"",ruta+"bluefireback2.gif"],150,0);
	stm_bp("p1",[1,4,0,0,5,5,0,0,85,"",-2,"",-2,50,2,3,"#999999","#333333","",0,1,1,"#498bfb"]);
		stm_aix("p1i0","p0i0",[0,"Directorio","","",-1,-1,0,"DirectorioNS.asp","_self","","","","",0,0,0,"","",0,0,0,0,1,"#00CCFF",1,"#498bfb",0,"","",3,0,0,0,"#FFFFF7","#000000","#FFFFFF","#000000","8pt Verdana","8pt Verdana"]);
		stm_aix("p2i1","p1i0",[0,"Balance General","","",-1,-1,0,"BalanceGeneralNS.asp"]);
		stm_aix("p2i2","p1i0",[0,"Estado Ganancias y Perdidas","","",-1,-1,0,"EstGananyPerdidasNS.asp"]);
		stm_aix("p2i3","p1i0",[0,"Estado Flujo Efectivo","","",-1,-1,0,"EstFlujoEfectivoNS.asp"]);
		stm_aix("p2i4","p1i0",[0,"Estado Cambio Patrimonio Neto","","",-1,-1,0,"CambioPatNetoNS.asp"]);
		stm_aix("p2i4","p1i0",[0,"Dividendos","","",-1,-1,0,"DividendosNS.asp"]);
		stm_aix("p2i5","p1i0",[0,"Consolidado EEFF","","",-1,-1,0,"ConsolidadoEEFFNS.asp"]);

stm_ep();

stm_aix("p0i3","p0i0",[0,"Salir","","",-1,-1,0,"logoffce.asp","_self","","",ruta+"gnome_logout_ov.png",ruta+"gnome_logout.png",16],90,0);
stm_ep();
stm_em();
