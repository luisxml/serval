_SCREEN.VISIBLE = .F.
_SCREEN.BorderStyle= 3
application.Visible = .F.

DECLARE INTEGER FindWindow IN WIN32API AS WAPIFWindow STRING @,STRING @
_aclass=NULL
_aname= [Inicio de Sesi�n - SERVAL]
_XAppwhd=WAPIFWindow( @_aclass, @_aname )
CLEAR DLLS
IF _XAppwhd!=0
	SET PROCEDURE TO
	SET LIBRARY TO
	MESSAGEBOX( [Disculpe, el sistema ya est� abierto.], 16, [Mensaje del Sistema] )
	CLOSE ALL
	CLEAR event
	CLOSE DATABASES
	RETURN
ENDIF

DECLARE INTEGER FindWindow IN WIN32API AS WAPIFWindow STRING @,STRING @
_aclass=NULL
_aname= [SERVAL - .:SISTEMA DE GESTION.:]
_XAppwhd=WAPIFWindow( @_aclass, @_aname )
CLEAR DLLS
IF _XAppwhd!=0
	SET PROCEDURE TO
	SET LIBRARY TO
	MESSAGEBOX( [Disculpe, el sistema ya est� abierto.], 16, [Mensaje del Sistema] )
	CLOSE ALL
	CLEAR event
	CLOSE DATABASES
	RETURN
ENDIF
RELEASE _aclass
*!*	*/////////////////////////////////////////////////////////////////////////////////
*!*	*// CHEQUEO DE EXCLUSIVIDAD DE INSTANCIA!!
*!*	*/////////////////////////////////////////////////////////////////////////////////
*!*	* basicamente lo que hace es buscar una ventana que tenga el mismo titulo y si lo encuentra
*!*	* significa que otra instancia ya se cargo previamente, hay una funcion API no recuerdo bien
*!*	* el nombre creo que se llamaba BingOnTop que te trae la ventana cuando esta no esta
*!*	* seleccionada
*!*	DECLARE INTEGER FindWindow IN WIN32API AS WAPIFWindow STRING @,STRING @
*!*	_aclass=NULL
*!*	_aname= [Inicio SBS Romanas]
*!*	_XAppwhd=WAPIFWindow( @_aclass, @_aname )
*!*	CLEAR DLLS
*!*	IF _XAppwhd!=0
*!*		SET PROCEDURE TO
*!*		SET LIBRARY TO
*!*		MESSAGEBOX( [Disculpe, el sistema ya est� abierto.], 16, [Mensaje del Sistema] )
*!*		CLOSE ALL
*!*		CLEAR event
*!*		CLOSE DATABASES
*!*		RETURN
*!*	ENDIF

*!*	DECLARE INTEGER FindWindow IN WIN32API AS WAPIFWindow STRING @,STRING @
*!*	_aclass=NULL
*!*	_aname= [SBS - Basculas]
*!*	_XAppwhd=WAPIFWindow( @_aclass, @_aname )
*!*	CLEAR DLLS
*!*	IF _XAppwhd!=0
*!*		SET PROCEDURE TO
*!*		SET LIBRARY TO
*!*		MESSAGEBOX( [Disculpe, el sistema ya est� abierto.], 16, [Mensaje del Sistema] )
*!*		CLOSE ALL
*!*		CLEAR event
*!*		CLOSE DATABASES
*!*		RETURN
*!*	ENDIF

*!*	RELEASE _aclass
*!*	*/////////////////////////////////////////////////////////////////////////////////

*!*	CLEAR ALL
*!*	CLOSE ALL
*!*	SET TALK OFF
*!*	set date to DMY
*!*	SET DEBUG OFF
*!*	SET ESCAPE OFF
*!*	SET HELP OFF
*!*	SET CENTURY ON
*!*	SET SAFETY OFF
*!*	SET DELETE ON
*!*	SET EXCLUSIVE OFF
*!*	SET SYSMENU OFF
*!*	SET BELL ON
*!*	**----------
*!*	SET RESOURCE off 
*!*	SET REFRESH TO 1,1 
*!*	SET EXCLUSIVE OFF 
*!*	SET UNIQUE OFF 
*!*	SET AUTOSAVE ON 
*!*	SET OPTIMIZE ON 
*!*	SET REPROCESS TO AUTOMATIC 
*!*	SET MULTILOCKS ON 


MODIFY WINDOWS SCREEN TITLE ".:: SERVAL - SISTEMA DE GESTION ::."


PUBLIC lcAppDir,Vgc_almatrab,Vgc_usuario,nres,Vpc_server,Vgc_impresora,Vgc_impresora2,Vgc_version 
PUBLIC Vlc_numbl,Vlc_numcont,Vlc_nombrebuque,Vlc_numviaje,Vlc_fechaatr,Vgc_cliente2,Vgc_cambiom,Vpc_nuevafecha,Vpn_fecha_desde2,Vpn_fecha_hasta2,Vgc_clic,Vgn_opt,Vgc_cliente2,conex,Vgn_tipousu,lcConnect
PUBLIC Vgc_clic,Vgn_opt,Vgc_cliente2,conex,Vgn_tipousu,Vgc_serie,Vgc_nomb_usu,Vgn_actual,fg_conectado,Vgn_super_user,VPC_llave,Vgc_caja,caja,Vgc_barras
PUBLIC Vgc_fg_barra_menu,VGC_LLAVE,Vgc_localmente
Vgc_localmente = 1
Vgc_barras =0 
Vgn_super_user=0
fg_conectado=0
Vgn_actual=0
Vgc_cliente2=""
Vgc_serie=""
conex=0
Vgn_tipousu=0
lcAppDir = upper(ADDBS(SYS(5) + SYS(2003)))
Vgc_almatrab=0
Vgc_nomb_usu=""
Vgc_usuario=""
Vgc_clic =0
Vgn_opt=1
conex=0
Vgc_fg_barra_menu = 0
VGC_LLAVE =''
*
*---------------------

*****VERSION
Vgc_version='2.1.4'
Vgc_vigencia='26/10/2020'
***VERSION
SET STEP ON
Local cnControladorArch,nTama�o,cCadena
cnControladorArch = FOPEN(lcAppDir+"config.txt")
*cnControladorArch = FOPEN("c:\config.txt")
nTama�o =  FSEEK(cnControladorArch, 0, 2)    && Lleva el puntero a EOF.
IF nTama�o <= 0
	MESSAGEBOX("Este archivo est� vac�o.")
ELSE
  = FSEEK(cnControladorArch, 0, 0)     && Mueve el puntero a BOF.
cCadena = FREAD(cnControladorArch, nTama�o)
ENDIF
= FCLOSE(cnControladorArch)   

aux=AT("$",cCadena,1)
Vgc_server=SUBSTR(cCadena,1,aux-1)
aux2=AT("$",cCadena,2)
Vgc_bd=SUBSTR(cCadena,aux+1,aux2-aux-1)
caja=ALLTRIM(SUBSTR(cCadena,aux2+1,LEN(cCadena)-aux2))
Vgc_caja=VAL(caja)
*---------------------

lcparServidor=Vgc_server
lcparDataBase =Vgc_bd
Vcl_conex = "driver={SQL Server};"+; 
"server=" + lcparServidor + ";"+; 
"database="+lcparDataBase+";"+;
" UID=serval; pwd=Cybers@c1;trusted_connection=no;" 
Conex=SQLSTRINGCONNECT(Vcl_conex)
*" UID=sa; pwd=sa123;trusted_connection=no;" 
**ERROR CONEXION SERVER
IF Conex > 0 THEN 
		
		SET DEFAULT TO FULLPATH(lcAppDir)
		*OPEN DATABASE FULLPATH(lcAppDir+"data\bppc.dbc")
		SET PATH TO "data,formularios,ing,ico,botones,clase,prg,reportes,archivos,menu"

****DATOS CONFIGURACION
lsql="Select * FROM  CONFIGURACION"
resp=SQLEXEC(conex, lsql, "CONFIGURACION")
IF resp<0
	MESSAGEBOX("Disculpe, error en la consulta, por favor comunicarse con el personal de soporte tecnico.",0+16,"Error de conexi�n")
	RETURN 
ELSE	
	Vpc_server = ALLTRIM(NB_SERVIDOR)	
	Vgc_impresora = ALLTRIM(NB_IMPRESORA)
	Vgc_impresora2 = ALLTRIM(NB_IMPRESORA2)
	*Vgc_localmente = ARCHIVOS_LOCALMENTE
ENDIF 
**** FIN DATOS CONFIGURACION

	****VERIFICAR EL SERIAL DEL DISCO
	Vlc_disco = LEFT( lcAppDir, 2)
	loFSO = CREATEOBJECT("Scripting.FileSystemObject")
	lcSerialNumber = lofso.drives(Vlc_disco).serialnumber 
	*messagebox((lcSerialNumber))
	*****FIN VERIFICAR SERIAL DEL DISCO
					
	****NOMBRE PC	
	Vlc_nom_pc=SUBSTR(SYS(0),1,AT('#',SYS(0))-2)
	lsql="select dbo.fn_encripta(?Vlc_nom_pc) as pc_encriptada"
	resp=SQLEXEC(conex,lsql,"llave")
	SELECT llave
	Vlc_pc_encriptada=ALLTRIM(pc_encriptada)
	****FIN NOMBRE PC

	****NOMBRE MAC
	LOCAL lcComputerName, loWMIService, loItems, loItem, lcMACAddress,Vlc_mac
	Vpc_mac_a=''
	lcComputerName = "."
	loWMIService = GETOBJECT("winmgmts:\\" + lcComputerName + "\root\cimv2")
	loItems = loWMIService.ExecQuery("Select * from Win32_NetworkAdapter",,48)

	FOR EACH loItem IN loItems
	lcMACAddress = loItem.MACAddress
		IF !ISNULL(lcMACAddress)		
			IF EMPTY(Vpc_mac_a)
				Vpc_mac_a = ALLTRIM(UPPER(loItem.MACAddress))
				*MESSAGEBOX(Vpc_mac_a)
			ENDIF   
			
		ENDIF
	ENDFOR
	lsql="select dbo.fn_encripta(?Vpc_mac_a) as mac_encriptada"
	resp=SQLEXEC(conex,lsql,"llave")
	SELECT llave
	Vlc_mac_encriptada=ALLTRIM(mac_encriptada)
	***FIN NOMBRE MAC

	*****VERIFICO KEY EN ARCHIVO TXT
	Local cnControladorArch2,nTama�o2,cCadena2
	cnControladorArch2 = FOPEN(lcAppDir+"data\temp\temp130713\temp.txt")
	nTama�o2 =  FSEEK(cnControladorArch2, 0, 2)    && Lleva el puntero a EOF.
	IF nTama�o2 <= 0
		*MESSAGEBOX("Este archivo est� vac�o.")
		cCadena2 =''
	ELSE
	  = FSEEK(cnControladorArch2, 0, 0)     && Mueve el puntero a BOF.
	cCadena2 = FREAD(cnControladorArch2, nTama�o2)
	VPC_llave = cCadena2
	ENDIF
	= FCLOSE(cnControladorArch2)   
	*********FIN VERIFICO KEY EN ARCHIVO TXT	


	lsql="select dbo.fn_desencripta(?cCadena2) as llave_desencriptada"
	resp=SQLEXEC(conex,lsql,"llave_des")
	SELECT llave_des
	Vlc_llave_desencriptada=ALLTRIM(llave_desencriptada)
	
	lsql="select * from NOTAS_REPORTES"
	resp=SQLEXEC(conex, lsql, "NOTAS_REPORTES")
	IF resp>0
		SET EXACT ON 		
		IF nTama�o2 > 0 	
			SELECT NOTAS_REPORTES
			
			*LOCATE FOR NOTA_1 = ALLTRIM(Vlc_llave_desencriptada) AND  NOTA_2 = ALLTRIM(Vlc_mac_encriptada) AND active = .t.
			LOCATE FOR NOTA_1 = ALLTRIM(Vlc_llave_desencriptada) AND  REPORTE_1 = (lcSerialNumber) AND active = .t.
		
			IF FOUND()								
				USE IN NOTAS_REPORTES
				USE IN LLAVE_DES					
				DO FORM inicio_venta
				READ EVENTS	
				*MESSAGEBOX('listo')			
			ELSE	
				MESSAGEBOX("Disculpe, Error en LICENCIA, por favor comunicarse con el Soporte Tecnico del Sistema .",0+16,"Error de Licencia")
				RETURN
				**DO FORM INICIO_venta
				**READ EVENTS
			ENDIF
		ELSE
			SELECT NOTAS_REPORTES
			*LOCATE FOR NOTA_2 = ALLTRIM(Vlc_mac_encriptada) AND active = .t.
			LOCATE FOR REPORTE_1 = (lcSerialNumber) AND active = .t.
			
			IF FOUND()
				Vlc_NOTA_1 = ALLTRIM(NOTA_1)
				lsql="select dbo.fn_encripta(?Vlc_NOTA_1) as llave_encriptada"
				resp=SQLEXEC(conex,lsql,"llave")
				SELECT llave
				Vlc_llave_encriptada=ALLTRIM(llave_encriptada)				
				
				lcFile=lcAppDir+"data\temp\temp130713\temp.txt"
				*lcCadena=(STR(lcSerialNumber))+'&'+ALLTRIM(Vlc_NOTA_1)
				lcCadena=ALLTRIM(Vlc_llave_encriptada)			
				Strtofile( lcCadena, lcFile ,1)
												
				lsql = "UPDATE NOTAS_REPORTES SET REPORTE_1= ?lcSerialNumber, STATUS = 1, TOTAL_REPORTE = ?Vlc_pc_encriptada, NOTA_2=?Vlc_mac_encriptada WHERE NOTA_1 = ?Vlc_NOTA_1"
				resp=SQLEXEC(conex, lsql)
				IF resp<0
					MESSAGEBOX("Disculpe, error en la consulta, por favor comunicarse con el Soporte Tecnico del Sistema .",0+16,"Error de conexi�n")
					RETURN 
				ENDIF
				USE IN NOTAS_REPORTES
				USE IN LLAVE_DES					
				DO FORM INICIO_venta
				READ EVENTS							
			ELSE
				SELECT NOTAS_REPORTES
				LOCATE FOR STATUS = 0 AND active = .t.
				IF FOUND()
					Vlc_NOTA_1 = ALLTRIM(NOTA_1)
					lsql="select dbo.fn_encripta(?Vlc_NOTA_1) as llave_encriptada"
					resp=SQLEXEC(conex,lsql,"llave")
					SELECT llave
					Vlc_llave_encriptada=ALLTRIM(llave_encriptada)				
					
					lcFile=lcAppDir+"data\temp\temp130713\temp.txt"
					*lcCadena=(STR(lcSerialNumber))+'&'+ALLTRIM(Vlc_NOTA_1)
					lcCadena=ALLTRIM(Vlc_llave_encriptada)			
					Strtofile( lcCadena, lcFile ,1)
													
					lsql = "UPDATE NOTAS_REPORTES SET REPORTE_1= ?lcSerialNumber, STATUS = 1, TOTAL_REPORTE = ?Vlc_pc_encriptada, NOTA_2=?Vlc_mac_encriptada,FH_ACTIVO=getdate() WHERE NOTA_1 = ?Vlc_NOTA_1"
					resp=SQLEXEC(conex, lsql)
					IF resp<0
						MESSAGEBOX("Disculpe, error en la consulta, por favor comunicarse con el Soporte Tecnico del Sistema .",0+16,"Error de conexi�n")
						RETURN 
					ENDIF
					USE IN NOTAS_REPORTES
					USE IN LLAVE_DES					
					DO FORM INICIO_venta
					READ EVENTS	
				ELSE
					MESSAGEBOX("Disculpe, Error en LICENCIA, por favor comunicarse con el Soporte Tecnico del Sistema .",0+16,"Error de Licencia")
					RETURN			
					**DO FORM INICIO_venta
					**READ EVENTS					
				ENDIF				
			ENDIF 		
		ENDIF 
	ELSE
		MESSAGEBOX("Disculpe, error en la consulta por favor comunicarse con el Soporte Tecnico del Sistema .",0+16,"Error de conexi�n")
		RETURN 
	ENDIF 		
ELSE
	MESSAGEBOX("Disculpe, error de conexion con el servidor de base de datos Soporte Tecnico del Sistema .",0+16,"Error de conexi�n")
	RETURN 
ENDIF 
************FIN ERROR CONEXION SERVER

*DO FORM actualizacion
*READ EVENTS	









