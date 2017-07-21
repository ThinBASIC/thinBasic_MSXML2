<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
   <LowPriceBooks>
      <xsl:for-each select="*/BOOK[not(PRICE >'220')]">
   <xsl:copy>
   <xsl:apply-templates select="*"/>
   </xsl:copy>
      </xsl:for-each>
   </LowPriceBooks>
  </xsl:template>
  <xsl:template match="*">
   <xsl:copy>
   <xsl:apply-templates />
   </xsl:copy>
  </xsl:template>
</xsl:stylesheet>
