<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">


  <xsl:template match="/">
    <xsl:text disable-output-escaping="yes"><![CDATA[<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">]]></xsl:text>

    <html>
      <body>
        <div class="container">
          <h1> Test Results </h1>
          <div>
            Tests run: <a href="#all">
              <b>
                <xsl:value-of select="sum(//assembly/@total)"/>
              </b>
            </a>;
            <xsl:if test="sum(//assembly/@errors) > 0">
              Errors: <a href="#errors">
                <b>
                  <xsl:value-of select="sum(//assembly/@errors)"/>
                </b>
              </a>,
            </xsl:if>
            <xsl:if test="sum(//assembly/@failed) > 0">
              Failures: <a href="#failures">
                <b>
                  <xsl:value-of select="sum(//assembly/@failed)"/>
                </b>
              </a>,
            </xsl:if>
            <xsl:if test="sum(//assembly/@skipped) > 0">
              Skipped: <a href="#skipped">
                <b>
                  <xsl:value-of select="sum(//assembly/@skipped)"/>
                </b>
              </a>,
            </xsl:if>
            Run time: <b>
              <xsl:value-of select="format-number(sum(//assembly/@time), '0.000')"/>s
            </b>
          </div>
          
          <table class="table table-striped" border="1" >
            <tr bgcolor="#ccccff">
              <th style="text-align:left">Test</th>
              <th style="text-align:left">Test Output</th>
              <th style="text-align:left">Errors</th>
              <th style="text-align:left">Results</th>
              <th style="text-align:left">time(s)</th>
            </tr>
            <xsl:for-each select="assemblies/assembly/class/test">
              <xsl:if test="@result = 'Pass'">
                <tr class="success">
                  <td>
                    <xsl:value-of select="@name"/>
                  </td>
                  <td>
                    <xsl:value-of select="output" disable-output-escaping="yes"/>
                  </td>
                  <td>No Errors</td>
                  <td>
                    <xsl:value-of select="@result"/>
                  </td>
                  <td>
                    <xsl:value-of select="@time"/>
                  </td>

                </tr>
              </xsl:if>
              <xsl:if test="@result = 'Fail'">
                <tr class="danger">
                  <td>
                    <xsl:value-of select="@name"/>
                  </td>
                  <td>
                    <td><xsl:value-of select="failure/message"/></td>
                  </td>
                  <td>
                    <xsl:value-of select="@result"/>
                  </td>
                  <td>
                    <xsl:value-of select="@time"/>
                  </td>
                </tr>
              </xsl:if>
            </xsl:for-each>
          </table>
        </div>
      </body>
    </html>
  </xsl:template>

</xsl:stylesheet>

