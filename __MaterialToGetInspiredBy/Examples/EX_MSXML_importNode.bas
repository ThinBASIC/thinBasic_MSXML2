' ========================================================================================
' Demonstrates the use of the importNode method.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL FreeThreadedDOMDocument AS IXMLDOMDocument3
   LOCAL pDOMDocument AS IXMLDOMDocument3
   LOCAL pIXMLDOMNode AS IXMLDOMNode
   LOCAL pCloneNode AS IXMLDOMNode
   LOCAL bstrMsg AS WSTRING
   LOCAL pElement AS IXMLDOMElement
   LOCAL pTextNode AS IXMLDOMNode

   FreeThreadedDOMDocument = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(FreeThreadedDOMDocument) THEN EXIT FUNCTION

   pDOMDocument = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pDOMDocument) THEN EXIT FUNCTION

   pDOMDocument.async = %VARIANT_FALSE
   IF ISFALSE pDOMDocument.Load("doc1.xml") THEN
      AfxShowMsg "Can't load doc1.xml"
      EXIT FUNCTION
   END IF

   FreeThreadedDOMDocument.async = %VARIANT_FALSE
   IF ISFALSE FreeThreadedDOMDocument.Load("doc2.xml") THEN
      AfxShowMsg "Can't load doc2.xml"
      EXIT FUNCTION
   END IF

   ' Copy a node from FreeThreadedDOMDocument to pDOMDocument:
   '   Fetch the "/doc" (node) from FreeThreadedDOMDocument (doc2.xml).
   '   Clone node for import to pDOMDocument.
   '   Append clone to pDOMDocument (doc1.xml).

   pIXMLDOMNode = FreeThreadedDOMDocument.selectSingleNode("/doc")
   pCloneNode = pDOMDocument.importNode(pIXMLDOMNode, %VARIANT_TRUE)
   pElement = pDOMDocument.documentElement
   pElement.appendChild pCloneNode
   pTextNode = pDOMDocument.createTextNode($CRLF)
   pElement.appendChild pTextNode
   pElement = NOTHING
   pIXMLDOMNode = NOTHING
   pCloneNode = NOTHING

   bstrMsg = bstrMsg + "doc1.xml after importing /doc from doc2.xml:"
   bstrMsg = bstrMsg + $CRLF & pDOMDocument.xml & $CRLF

   ' Clone a node using importNode() and append it to self:
   '   Fetch the "doc/b" (node) from pDOMDocument (doc1.xml).
   '   Clone node using importNode() on pDOMDocument.
   '   Append clone to pDOMDocument (doc1.xml).

   pIXMLDOMNode = FreeThreadedDOMDocument.selectSingleNode("/doc/b")
   pCloneNode = pDOMDocument.importNode(pIXMLDOMNode, %VARIANT_TRUE)
   pElement = pDOMDocument.documentElement
   pElement.appendChild pCloneNode
   pTextNode = pDOMDocument.createTextNode($TAB)
   pElement.appendChild pTextNode
   pElement = NOTHING
   pIXMLDOMNode = NOTHING
   pCloneNode = NOTHING

   bstrMsg = bstrMsg & "doc1.xml after import /doc/b from self:"
   bstrMsg = bstrMsg & $CRLF & pDOMDocument.xml & $CRLF

   ' Clone a node and append it to the dom using cloneNode():
   '   Fetch "doc/a" (node) from pDOMDocument (doc1.xml).
   '   Clone node using cloneNode on pDOMDocument.
   '   Append clone to pDOMDocument (doc1.xml).

   pIXMLDOMNode = FreeThreadedDOMDocument.selectSingleNode("/doc/a")
   pCloneNode = pDOMDocument.importNode(pIXMLDOMNode, %VARIANT_TRUE)
   pElement = pDOMDocument.documentElement
   pElement.appendChild pCloneNode
   pTextNode = pDOMDocument.createTextNode($TAB)
   pElement.appendChild pTextNode
   pElement = NOTHING
   pIXMLDOMNode = NOTHING
   pCloneNode = NOTHING

   bstrMsg = bstrMsg & "doc1.xml after cloning /doc/a from self:"
   bstrMsg = bstrMsg & $CRLF + pDOMDocument.xml & $CRLF

   pDOMDocument.save "out.xml"
   bstrMsg = bstrMsg + "a new document was saved to out.xml in the current working directory."
   AfxShowMsg bstrMsg

END FUNCTION
' ========================================================================================
