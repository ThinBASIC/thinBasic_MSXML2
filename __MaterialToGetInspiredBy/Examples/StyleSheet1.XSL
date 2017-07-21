<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
   <PriceList>
      <xsl:for-each select="COLLECTION/BOOK">
   <xsl:sort select="TITLE" data-type="text"/>
   <xsl:copy>
   <xsl:apply-templates select="*"/>
   </xsl:copy>
      </xsl:for-each>
   </PriceList>
  </xsl:template>

  <xsl:template match="*">
   <xsl:copy>
   <xsl:apply-templates />
   </xsl:copy>
  </xsl:template>
</xsl:stylesheet>
