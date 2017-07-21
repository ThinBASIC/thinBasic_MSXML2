' ========================================================================================
' Demonstrates the use of the selectNodes method.
' The following example creates an IXMLDOMNodeList object containing the nodes specified
' by the expression parameter (for example, all the <xsl:template> nodes in an XSLT style
' sheet). It then displays the number of nodes contained in the node list.
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
   LOCAL pNodeList AS IXMLDOMNodeList

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "hello.xsl"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pXmlDoc.setProperty "SelectionNamespaces", "xmlns:xsl='http://www.w3.org/1999/XSL/Transform'"
         pXmlDOc.setProperty "SelectionLanguage", "XPath"
         pNodeList = pXmlDoc.documentElement.selectNodes("//xsl:template")
         AfxShowMsg FORMAT$(pNodeList.length)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
