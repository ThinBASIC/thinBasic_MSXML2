' ========================================================================================
' Demonstrates the use of the insertBefore method.
' The following example creates a new IXMLDOMNode object and inserts it as the second
' child of the top-level node.
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
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pItem AS IXMLDOMNode
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pNewNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.loadXML "<root><child1/></root>"
   pRootNode = pXmlDoc.documentElement
   AfxShowMsg pRootNode.xml
   pNewNode = pXmlDoc.createNode(%NODE_ELEMENT, "CHILD2", "")
   pNodeList = pRootNode.childNodes
   pItem = pNodeList.item(0)
   pCurrNode = pRootNode.insertBefore(pNewNode, pItem)
   AfxShowMsg pRootNode.xml

END FUNCTION
' ========================================================================================
