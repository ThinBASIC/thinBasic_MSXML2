' ========================================================================================
' Demonstrates the use of the getOption method.
' This example returns the current value of the %SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
' (option 2), which by default is 13056. This value maps to the
' %SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS flag, indicating that the current XMLHTTP
' server instance will return all certificate errors.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlServerHttp AS IServerXMLHTTPRequest
   LOCAL vValue AS VARIANT

   pXmlServerHttp = NEWCOM "Msxml2.ServerXMLHTTP.6.0"
   IF ISNOTHING(pXmlServerHttp) THEN EXIT FUNCTION

   vValue = pXmlServerHttp.getOption(2)
   AfxShowMsg STR$(VARIANT#(vValue))

END FUNCTION
' ========================================================================================
