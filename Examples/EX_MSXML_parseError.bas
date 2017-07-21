' ========================================================================================
' Demonstrates the use of the XMLDOMParseError interface.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument2
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL bstrMsg AS STRING

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISFALSE ISOBJECT(pXmlDoc) THEN EXIT FUNCTION

   IF ISFALSE pXmlDoc.load("bad.xml") THEN
      pParseError = pXmlDoc.parseError
      IF ISTRUE ISOBJECT(pParseError) THEN
         bstrMsg  = "-------------------------------------------------------------------------" & $CRLF
         bstrMsg += "Error &H" & HEX$(pParseError.errorCode) & " in the xml file" & $CRLF
         bstrMsg += "-------------------------------------------------------------------------" & $CRLF
         bstrMsg += "Filepos:     " & FORMAT$(pParseError.filePos) & $CRLF
         bstrMsg += "Line:        " & FORMAT$(pParseError.line) & $CRLF
         bstrMsg += "Position:    " & FORMAT$(pParseError.linePos) & $CRLF
         bstrMsg += "Reason:      " & pParseError.reason & $CRLF
         bstrMsg += "Source text: " & pParseError.srcText & $CRLF
         bstrMsg += "Url:         " & pParseError.url & $CRLF
         bstrMsg += "-------------------------------------------------------------------------" & $CRLF
         pParseError = NOTHING
         AfxShowMsg bstrMsg
      END IF
   END IF

   pXmlDoc = NOTHING

END FUNCTION
' ========================================================================================
