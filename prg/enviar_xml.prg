*FUNCTION FE_SUNATWSENVIO


 PARAMETER _rutaxml, _nombrearchivose, femodo, rucsol, usuariosol, clavesol, purlws,ruta_cdr_local,ruta_cdr_server,ruta_rpta_server&& ejm: "c:\", "20441053658-01-F050-0000001" , "P", "xxxxxxxxxxx","xxx","xxx", "direccion del servicio"

SET STEP ON
 ls_ruc_emisor = rucsol
 ls_user = ls_ruc_emisor+usuariosol
 ls_pwd_sol = clavesol
 ps_file = _rutaxml+_nombrearchivose
 ps_filezip = ps_file+".zip"
 IF  .NOT. FILE(ps_filezip)
    *RETURN "Archivo zip a enviar no encontrado"
    RETURN 0
 ENDIF
 ls_filename = JUSTFNAME(ps_filezip)
 ls_contentfile = FILETOSTR(ps_filezip)
 ls_base64 = STRCONV(ls_contentfile, 13)
 TEXT TO ls_envioxml TEXTMERGE NOSHOW PRETEXT 0015 FLAGS 1
			<soapenv:Envelope xmlns:ser="http://service.sunat.gob.pe" xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
					xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
				<soapenv:Header>
					<wsse:Security>
						<wsse:UsernameToken>
							<wsse:Username><<ls_user>></wsse:Username>
							<wsse:Password><<ls_pwd_sol>></wsse:Password>
						</wsse:UsernameToken>
					</wsse:Security>
				</soapenv:Header>
				<soapenv:Body>
					<ser:sendBill>
						<fileName><<ls_fileName>></fileName>
						<contentFile><<ls_base64>></contentFile>
					</ser:sendBill>
				</soapenv:Body>				
			</soapenv:Envelope>
 ENDTEXT
 oxmlhttp = CREATEOBJECT("MSXML2.ServerXMLHTTP.6.0")
 oxmlbody = CREATEOBJECT('MSXML2.DOMDocument.6.0')
 IF  .NOT. (oxmlbody.loadxml(ls_envioxml))
    oresp.mensaje = "No se cargo XML: "+oxmlbody.parseerror.reason
    *RETURN .F.
    RETURN "<message>ERROR NO SE CARGO XML</message>"
 ENDIF
 lsurl = purlws
 oxmlhttp.open('POST', lsurl, .F.)
 oxmlhttp.setrequestheader("Content-Type", "text/xml")
 oxmlhttp.setrequestheader("Content-Type", "text/xml;charset=ISO-8859-1")
 oxmlhttp.setrequestheader("Content-Length", LEN(ls_envioxml))
 oxmlhttp.setrequestheader("SOAPAction", "sendBill")
 oxmlhttp.setoption(2, 13056) 
 
 lcErr = "" 
TRY 
 oxmlhttp.send(oxmlbody.documentelement.xml)

CATCH TO loError 
	lcErr = [Error: ] + STR(loError.ERRORNO) + CHR(13) + [Linea: ] + STR(loError.LINENO) + CHR(13) + [Mensaje: ] + loError.MESSAGE 
FINALLY 
	IF EMPTY(lcErr) 
		*MESSAGEBOX("El mensaje se envió con éxito", 64, "Aviso") 
	ELSE 
		MESSAGEBOX(lcErr, 16 , "Error")
	ENDIF 
ENDTRY 

IF EMPTY(lcerr)
ELSE
	*RETURN "ERROR- ASEGURE QUE HAY CONEXION A INTENET O ESTE DISPONIBLE EL SERVIDOR EL COMPROBANTE NO SE ENVIO" 
	*RETURN 0
	RETURN "ERROR - NO QUE HAY CONEXION A INTENET O  EL SERVIDOR NO ESTA DISPONIBLE, EL <message>COMPROBANTE NO SE ENVIO</message> <httpstatus>"+"ERROR"+"</httpstatus>"&&+ +NVL(oxmlhttp.responsetext, '')
ENDIF

 
 SET STEP ON
 IF (oxmlhttp.status<>200)
 	*MESSAGEBOX("ERROR:"+ oxmlhttp.responsetext )
    RETURN "<message>Error: Comprobante NO SE HA ENVIADO</message>, ERROR: <httpstatus>"+ALLTRIM(STR(oxmlhttp.status))+"</httpstatus>"+ +NVL(oxmlhttp.responsetext, '')
    *RETURN 0
 ELSE
    loxmlresp = CREATEOBJECT("MSXML2.DOMDocument.6.0")
    loxmlresp.loadxml(oxmlhttp.responsetext)
    ccontenidorptazip = STREXTRACT(oxmlhttp.responsetext, "<applicationResponse>", "</applicationResponse>")
    *MESSAGEBOX(oxmlhttp.responsetext)
    
    Vlc_faultcode = STREXTRACT(oxmlhttp.responsetext, "<faultcode>", "</faultcode>")
    Vlc_faultstring= STREXTRACT(oxmlhttp.responsetext, "<faultstring>", "</faultstring>")
    Vlc_message= STREXTRACT(oxmlhttp.responsetext, "<message>", "</message>")
    
    IF EMPTY(Vlc_faultcode) AND EMPTY(Vlc_faultstring) AND EMPTY(Vlc_message) THEN 
    
	    DELETE FILE ALLTRIM(ruta_cdr_local+"R-"+_nombrearchivose+".zip")
	    STRTOFILE(STRCONV(ccontenidorptazip, 14), ruta_cdr_local+"R-"+_nombrearchivose+".zip")
	    
	    DELETE FILE ALLTRIM(ruta_cdr_server+"R-"+_nombrearchivose+".zip")
	    STRTOFILE(STRCONV(ccontenidorptazip, 14), ruta_cdr_server+"R-"+_nombrearchivose+".zip")
	     
	    DELETE FILE ALLTRIM(ruta_rpta_server+"R-"+_nombrearchivose+".zip")
	    STRTOFILE(STRCONV(ccontenidorptazip, 14), ruta_rpta_server+"R-"+_nombrearchivose+".zip")
	    *STRTOFILE(STRCONV(ccontenidorptazip, 14), _rutaxml_cdr+"R-"+_nombrearchivose_CLIENTE+".zip")
	    *STRTOFILE(STRCONV(ccontenidorptazip, 14), _rutaxml_cdr_server+"R-"+_nombrearchivose_CLIENTE+".zip")  
	    *MESSAGEBOX(oxmlhttp.responsetext)
	    *RETURN "COMPROBANTE ENVIADO Y ACEPTADO"
	    RETURN '<message>ENVIADO Y ACEPTADO</message> '+'<httpstatus>'+ALLTRIM(STR(oxmlhttp.status))+'</httpstatus>'+'-'+NVL(oxmlhttp.responsetext, '')
	    
    ELSE
    	RETURN oxmlhttp.responsetext
    ENDIF 
    
    
    * rptawsticket = 
 ENDIF

*ENDFUNC
