package com.lowcodeminds.appian.plugins.excel;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

/**
 * Input CDT for {@link ExcelExportSmartService}'s {@code SQLSheetDataList} input: one
 * SQL query and the sheet name its results should be written to.
 */
@XmlRootElement(name = "SQLSheetData", namespace = "urn:com.lowcodeminds:types:SQLSheetData")
@XmlType(name = "SQLSheetData", namespace = "urn:com.lowcodeminds:types:SQLSheetData", propOrder = {"sheetName", "sqlQuery"})
public class SQLSheetData {

  private String sheetName;
  private String sqlQuery;

  @XmlElement
  public String getSheetName() {
    return sheetName;
  }

  public void setSheetName(String sheetName) {
    this.sheetName = sheetName;
  }

  @XmlElement
  public String getSqlQuery() {
    return sqlQuery;
  }

  public void setSqlQuery(String sqlQuery) {
    this.sqlQuery = sqlQuery;
  }
}
