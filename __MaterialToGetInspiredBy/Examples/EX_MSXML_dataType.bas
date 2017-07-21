' ========================================================================================
' Demonstrates the use of the dataType property.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXMLDoc AS IXMLDOMDocument
   LOCAL pRootNode AS IXMLDOMElement

   pXMLDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXMLDoc) THEN EXIT FUNCTION

   TRY
      pXMLDoc.async = %VARIANT_FALSE
      pXMLDoc.loadXML("<root/>")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXMLDoc.documentElement
         pRootNode.dataType = "int"
         pRootNode.nodeTypedValue = 5
         AfxShowMsg pXMLDoc.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
