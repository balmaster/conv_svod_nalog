<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:w="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:html="http://www.w3.org/TR/REC-html40"
    >
    <xd:doc xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl" scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>Created on:</xd:b> Aug 27, 2010</xd:p>
            <xd:p><xd:b>Author:</xd:b> balmaster</xd:p>
            <xd:p></xd:p>
        </xd:desc>
    </xd:doc>
    
    <xsl:template match="@*|node()">
           <xsl:apply-templates/>
    </xsl:template>

    <xsl:template match="w:Workbook">
        <xsl:element name="reports">
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>

    <xsl:template match="w:Worksheet[w:Table/w:Row/w:Cell/w:Data = 'А']">
        <xsl:value-of select="'&#xA;'"/>
        <xsl:element name="report">
            <xsl:attribute name="name">
                <xsl:value-of select="@ss:Name"/>
            </xsl:attribute>
            
            <xsl:variable name="cellCount" select="count(w:Table/w:Row[w:Cell/w:Data = 'А'][1]/w:Cell/w:Data)"/>
            <xsl:attribute name="count">
                <xsl:value-of select="$cellCount"/>
            </xsl:attribute>
                        
            <xsl:for-each select="w:Table/w:Row[not (w:Cell/w:Data = 'А')]">
                <xsl:value-of select="'&#xA;'"/>
                <xsl:element name="p">
                    <xsl:for-each select="w:Cell">
                        <xsl:variable name="curPos" select="position()"/>
                        <xsl:value-of select="'&#xA;'"/>
                        <xsl:variable name="type" select="../../w:Row[w:Cell/w:Data = 'А']/w:Cell[$curPos]/w:Data"/>
                        <xsl:variable name="value" select="w:Data"/>
                            <xsl:if test="$type and $value">
                            <xsl:element name="pv">
                                <xsl:attribute name="type">
                                    <xsl:value-of select="$type"/>
                                </xsl:attribute>
                                <xsl:attribute name="value">
                                    <xsl:value-of select="$value"/>
                                </xsl:attribute>
                            </xsl:element>
                        </xsl:if>
                    </xsl:for-each>
                    <!-- <xsl:value-of select="'&#xA;'"/>-->
                </xsl:element>
            </xsl:for-each>
        
        </xsl:element>
    </xsl:template>
    
    
</xsl:stylesheet>
