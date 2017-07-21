<?xml version='1.0'?>
<xsl:stylesheet version="1.0" 
     xmlns:ms="urn:schemas-microsoft-com:xslt"   
     xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:output method="html"   
     omit-xml-declaration="yes"/>

  <xsl:template match="/">
     <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="*">
    <DIV>
       <xsl:value-of select="name()"/> is of
       "<xsl:value-of select="ms:type-local-name()"/>" in 
       "<xsl:value-of select="ms:type-namespace-uri()"/>" 
    </DIV>
    <xsl:apply-templates/>
  </xsl:template>

</xsl:stylesheet>
