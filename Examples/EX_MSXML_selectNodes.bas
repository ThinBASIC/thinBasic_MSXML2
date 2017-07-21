' ========================================================================================
' Demonstrates the use of the selectNodes method.
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

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeList = pXmlDoc.documentElement.selectNodes("//book")
         AfxShowMsg STR$(pNodeList.length)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
