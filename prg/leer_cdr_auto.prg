PARAMETER archivo, destino, nombre_archivo
*SET PATH TO FULLPATH(CURDIR())
SET CLASSLIB TO lcAppDir+'clase\zip' ADDITIVE 
*SET CLASSLIB TO LOCFILE('zip.vcx')  ADDITIVE 
*LOCFILE("FoxBarcodeQR.prg")
oZip=CREATEOBJECT('Zip.Zip')



**DESCOMPRIMIR CDR
WITH oZip
	.ArchivoZip= archivo	&& Ruta y nombre del archivo ZIP
	.DirectorioDestino= destino	&& Directorio de destino del contenido del ZIP
	*.Contrase�a='contrase�a'	&& Contrase�a para desproteger el zip
	.Descomprimir()
ENDWITH
**FIN DESCONPRIMIR CDR


****VERIFICAR SI SE DESCARGO CDR
LOCAL nTama�o, Vln_contador  
nTama�o = 0
Vln_contador = 0
DO WHILE nTama�o <= 0
	Wait Window "Esperando respuesta de SUNAT Por Favor Espere..." Timeout 5
   	directorio_local = 	destino+"\"+nombre_archivo
   	cnControladorArch = ''
	IF FILE(directorio_local)	
		cnControladorArch = FOPEN(directorio_local)								
	ENDIF
			
	nTama�o =  FSEEK(cnControladorArch, 0, 2)    && Lleva el puntero a EOF.
	*MESSAGEBOX(nTama�o)
	IF nTama�o <= 0 &&AND Vln_contador = 1000
		*Vln_nres = MESSAGEBOX("�No se ha recibido ninguna respuesta, desea intentar de nuevo?",32+4,"Respuesta SUNAT")
		*IF Vln_nres <> 6
			*nTama�o =1
			 *EXIT 
			 RETURN "<CodigoResCDR>400</CodigoResCDR>"+"<ResponseCDR>Error en servicio SUNAT</ResponseCDR>"
		*ELSE
			*Vln_contador = 0
		*ENDIF 			 
		*IF Vln_nres= 6
		*Wait Window "Firmando Documento...Por Favor Espere..." Timeout 5					
	ENDIF 
	Vln_contador = Vln_contador + 1
ENDDO
****FIN VERIFICAR DESCARGA DE CDR
	    
****LEER RESPUESTA DE ARCHIVO CDR
IF nTama�o > 0			
	= FSEEK(cnControladorArch, 0, 0)     && Mueve el puntero a BOF.
	cCadena = FREAD(cnControladorArch, nTama�o)
	= FCLOSE(cnControladorArch)   
			
	Vlc_respuesta = STREXTRACT(cCadena, "<cbc:Description>", "</cbc:Description>")
	Vlc_codigo_respuesta = ALLTRIM(STREXTRACT(cCadena, "<cbc:ResponseCode>", "</cbc:ResponseCode>"))
	Vlc_codigo_respuesta2 = ALLTRIM(STREXTRACT(cCadena, '<cbc:ResponseCode listAgencyName="PE:SUNAT">', "</cbc:ResponseCode>"))
	Vlc_hash = STREXTRACT(cCadena, "<cbc:DocumentHash>", "</cbc:DocumentHash>")
	Vlc_observacion = STREXTRACT(cCadena, "<cbc:Note>", "</cbc:Note>")
	
	IF Vlc_codigo_respuesta = '0' OR Vlc_codigo_respuesta2 = '0'
		Vlc_codigo = '0'
		RETURN "<CodigoResCDR>"+Vlc_codigo+"</CodigoResCDR>"+"<ResponseCDR>"+ALLTRIM(Vlc_respuesta)+"</ResponseCDR>"
	ELSE
		Vlc_codigo = Vlc_codigo_respuesta + Vlc_codigo_respuesta2
		RETURN "<CodigoResCDR>"+Vlc_codigo+"</CodigoResCDR>"+"<ResponseCDR>"+ALLTRIM(Vlc_respuesta)+"</ResponseCDR>"
	ENDIF 
ENDIF 
****FIN LEER RESPUESTA DE ARCHIVO CDR
 
	
	
	