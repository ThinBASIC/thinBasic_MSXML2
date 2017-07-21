' ========================================================================================
' Demonstrates the use of fistChild property.
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
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pObjText AS IXMLDOMText
   LOCAL pFirstChild AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load("books.xml")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDOc.documentElement
         pObjText = pXmlDOc.createTextNode("Hello World!")
         pFirstChild = pRootNode.firstChild
         pRootNode.insertBefore pObjText, pFirstChild
         AfxShowMsg pRootNode.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
