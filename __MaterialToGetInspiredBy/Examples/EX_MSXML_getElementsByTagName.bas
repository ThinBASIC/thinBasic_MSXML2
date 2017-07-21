' ========================================================================================
' Demonstrates the use of the getElementsByTagName method.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL bstrOutput AS STRING
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISFALSE ISOBJECT(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmLDoc.load("books.xml")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeList = pXmlDoc.getElementsByTagName("author")
         FOR i = 0 TO pNodeList.length - 1
            bstrOutput += pNodeList.item(i).xml & $CRLF
         NEXT
         AfxShowMsg bstrOutput
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
