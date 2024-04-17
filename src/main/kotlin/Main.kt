package org.mvk

import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.jdom2.Document
import org.jdom2.input.SAXBuilder
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream


fun main() {
  val excelFile = File("src/main/resources/Connect_StringsTranslations_Consolidated.xlsx")
  val xmlFile = File("src/main/resources/strings.xml")

  val workbook = FileInputStream(excelFile).use { fileStream ->
    XSSFWorkbook(fileStream)
  }

  val xmlStrings = parseXmlFile(xmlFile)

  for (sheetIndex in 0 until workbook.numberOfSheets) {
    val sheet = workbook.getSheetAt(sheetIndex)
    processSheet(sheet, xmlStrings)
  }

  val outputStream = FileOutputStream(excelFile)
  workbook.write(outputStream)
  outputStream.close()
}

fun parseXmlFile(xmlFile: File): Map<String, String> {
  val saxBuilder = SAXBuilder()
  val document: Document = saxBuilder.build(xmlFile)
  val rootElement = document.rootElement
  val strings = rootElement.children.associate { element ->
    val name = element.getAttributeValue("name")
    val value = element.text
    // Key is the actual string and value is the ID
    value to name
  }
  return strings
}

fun processSheet(sheet: Sheet, xmlStrings: Map<String, String>) {
  val headerRow = sheet.getRow(0)
  val lastColumnNum = headerRow.lastCellNum.toInt()

  // Create a new cell ID if it doesn't exist at the end of all the columns
  if (headerRow.getCell(lastColumnNum)?.stringCellValue != "ID") {
    headerRow.createCell(lastColumnNum).setCellValue("ID")
  }

  // Iterate all rows except title
  for (rowIndex in 1..sheet.lastRowNum) {
    val row = sheet.getRow(rowIndex) ?: continue

    val cell = row.getCell(0)
    println("$rowIndex: $cell")
    val formatter = DataFormatter()
    try {
      val cellValue = cell.stringCellValue
      xmlStrings[cellValue]?.let { xmlName ->
        row.createCell(row.lastCellNum.toInt()).setCellValue(xmlName)
      }
    } catch (e: Exception) {
      println(e.message)
    } finally {
      val cellValue = formatter.formatCellValue(cell)
      xmlStrings[cellValue]?.let { xmlName ->
        row.createCell(lastColumnNum).setCellValue(xmlName)
      }
    }
  }
}
