<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  >

<xsl:output method="xml" version="1.0" encoding="utf-8" indent="yes" />


<!-- name of top level/root element -->
<xsl:param name="topLevelName" select="'TABLES'" />

<!-- name of an attribute of the top level element -->
<xsl:param name="topLevelAttribute" />

<!-- value of the attribute of the top level element -->
<xsl:param name="topLevelAttributeValue" />

<!-- name of the table row element -->
<xsl:param name="rowName" select="'ROW'" />

<!-- whitespace substitution (eg in tag names) -->
<xsl:param name="wsSub" select="'_'" />

<!-- line in which are the column headers -->
<xsl:param name="colNamesRow" select="1" />

<!-- line from which the data material begins -->
<xsl:param name="dataRowStart" select="2" />

<!-- table stuff -->
<xsl:variable name="content" select="/office:document-content/office:body/office:spreadsheet" />


<!-- ==================== default mode ==================== -->

<!-- entry template -->
<xsl:template match="/">
  <xsl:element name="{ translate( $topLevelName, '&#x20;&#x9;&#xA;', $wsSub ) }">
    
    <!-- optional attribute -->
    <xsl:if test="$topLevelAttribute">
      <xsl:attribute name="{ $topLevelAttribute }">
        <xsl:value-of select="$topLevelAttributeValue" />
      </xsl:attribute>
    </xsl:if>

    <!-- call tables -->
    <xsl:apply-templates select="$content/table:table" />
    
  </xsl:element>
</xsl:template>


<!-- table template -->
<xsl:template match="table:table">
  <xsl:element name="{ translate( @table:name, '&#x20;&#x9;&#xA;', $wsSub ) }">
    
    <!-- call all rows from the row that contains data -->
    <xsl:apply-templates select="table:table-row[ position() &gt;= $dataRowStart ]">
      <xsl:with-param name="colHeadlines" select="table:table-row[ position() = $colNamesRow ]/table:table-cell/text:p" />
    </xsl:apply-templates>
    
  </xsl:element>
</xsl:template>


<!-- row wrapper -->
<xsl:template match="table:table-row">
  <xsl:param name="colHeadlines" />
  <xsl:apply-templates mode="row" select=".">
    <xsl:with-param name="colHeadlines" select="$colHeadlines" />
  </xsl:apply-templates>
</xsl:template>


<!-- cell wrapper -->
<xsl:template match="table:table-cell">
  <xsl:param name="colHeadlines" />
  
  <!-- number of this cell -->
  <xsl:variable name="position" select="position()" />
  
  <!-- repetition of this cell -->
  <xsl:variable name="repeatings-here" select="
      concat (
        substring (
          @table:number-columns-repeated,
          1 div @table:number-columns-repeated ),
        substring (
          0,
          1 div not (@table:number-columns-repeated) )
      )"
    />
  
  <!-- repetitions before -->
  <xsl:variable name="repeatings-before" select="
      concat (
        substring (
          sum (preceding-sibling::table:table-cell[@table:number-columns-repeated]/@table:number-columns-repeated),
          1 div preceding-sibling::table:table-cell[@table:number-columns-repeated]/@table:number-columns-repeated ),
        substring (
          0, 
          1 div not (preceding-sibling::table:table-cell[@table:number-columns-repeated]/@table:number-columns-repeated) )
      )"
    />
  
  <!-- repetition blocks before -->
  <xsl:variable name="repeating-blocks-before" select="
      concat (
        substring (
          count (preceding-sibling::table:table-cell[@table:number-columns-repeated]/@table:number-columns-repeated),
          1 div preceding-sibling::table:table-cell[@table:number-columns-repeated]/@table:number-columns-repeated ),
        substring (
          0,
          1 div not (preceding-sibling::table:table-cell[@table:number-columns-repeated]/@table:number-columns-repeated) )
      )"
    />
  
  <!-- table header (depending on possible repetitions) -->
  <xsl:variable name="name"
    select="$colHeadlines[position () = $position + $repeatings-before - $repeating-blocks-before]"
    />
  
  <xsl:apply-templates mode="cell" select=".">
    <xsl:with-param name="colHeadlines"            select="$colHeadlines" />
    <xsl:with-param name="position"                select="$position" />
    <xsl:with-param name="repeatings-here"         select="$repeatings-here" />
    <xsl:with-param name="repeatings-before"       select="$repeatings-before" />
    <xsl:with-param name="repeating-blocks-before" select="$repeating-blocks-before" />
    <xsl:with-param name="name"                    select="$name" />
  </xsl:apply-templates>
</xsl:template>


<!-- ==================== mode="row" ==================== -->

<!-- row templates -->
<xsl:template mode="row" match=" node() | @* " />

<xsl:template mode="row" match="table:table-row[
        not ( @table:number-rows-repeated                                               )
        and ( table:table-cell[1]/@table:number-columns-repeated                &gt;99  )
    and not ( preceding-sibling::table:table-row[1][@table:number-rows-repeated &gt;99] )
    ]">
  <xsl:element name="{ translate( $rowName, '&#x20;&#x9;&#xA;', $wsSub ) }" />
</xsl:template>

<xsl:template mode="row" match="table:table-row[ @table:number-rows-repeated and ( table:table-cell[1]/@table:number-columns-repeated &gt;99 ) ]" />

<xsl:template mode="row" match="table:table-row[
    not ( @table:number-rows-repeated )
    and ( table:table-cell[1]/@table:number-columns-repeated   &gt;99 )
    and (
          preceding-sibling::table:table-row
          [1]
          [@table:number-rows-repeated                         &gt;99]
          /table:table-cell[@table:number-columns-repeated     &gt;99] )
    ]" />

<xsl:template mode="row" match="table:table-row[ (@table:number-rows-repeated &lt;99) and (table:table-cell[1]/@table:number-columns-repeated &gt;99) ]">
  <xsl:call-template name="emptyRows">
    <xsl:with-param name="counter" select="@table:number-rows-repeated" />
    <xsl:with-param name="step"    select="1" />
  </xsl:call-template>
</xsl:template>

<xsl:template mode="row" match="table:table-row">
  <xsl:param name="colHeadlines" />
  <xsl:element name="{ translate( $rowName, '&#x20;&#x9;&#xA;', $wsSub ) }">
    <xsl:apply-templates select="table:table-cell">
      <xsl:with-param name="colHeadlines" select="$colHeadlines" />
    </xsl:apply-templates>
  </xsl:element>
</xsl:template>


<!-- ==================== mode="cell" ==================== -->

<!-- cell templates -->
<xsl:template mode="cell" match=" node() | @* " />

<xsl:template mode="cell" match="table:table-cell[ @table:number-columns-repeated and (count( child::* ) =1) ]">
  <xsl:param name="colHeadlines" />
  <xsl:param name="position" />
  <xsl:param name="repeatings-here" />
  <xsl:param name="repeatings-before" />
  <xsl:param name="repeating-blocks-before" />
  <xsl:param name="name" />
  
  <xsl:variable name="value" select="normalize-space( text:p )" />
  
  <xsl:for-each select="$colHeadlines[
          ( position () &gt;= $position + $repeatings-before - $repeating-blocks-before )
      and ( position () &lt;  $position + $repeatings-before + $repeatings-here         )
    ]">
    <xsl:element name="{ translate( . , '&#x20;&#x9;&#xA;', $wsSub ) }">
      <xsl:value-of disable-output-escaping="yes" select="$value" />
    </xsl:element>
  </xsl:for-each>
</xsl:template>

<xsl:template mode="cell" match="table:table-cell[ not( @table:number-columns-repeated ) and ( count(child::*) =1 ) ]">
  <xsl:param name="name" />
  <xsl:element name="{ translate ($name, '&#x20;&#x9;&#xA;', $wsSub) }">
    <xsl:value-of disable-output-escaping="yes" select="normalize-space( text:p )" />
  </xsl:element>
</xsl:template>


<!-- ==================== named templates ==================== -->

<!-- template for emtpy rows -->
<xsl:template name="emptyRows">
  <xsl:param name="counter" />
  <xsl:param name="step" />
  
  <xsl:element name="{ translate( $rowName, '&#x20;&#x9;&#xA;', $wsSub ) }" />
  
  <xsl:if test="$step &lt; $counter">
    <xsl:call-template name="emptyRows">
      <xsl:with-param name="counter" select="$counter" />
      <xsl:with-param name="step"    select="$step +1" />
    </xsl:call-template>
  </xsl:if>
</xsl:template>


</xsl:stylesheet>
