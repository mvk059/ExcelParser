package org.mvk

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.jdom2.Document
import org.jdom2.Element
import org.jdom2.output.Format
import org.jdom2.output.XMLOutputter
import java.io.File
import java.io.FileOutputStream

fun main() {
  // Open the Excel file
  val workbook = WorkbookFactory.create(File("src/main/resources/Connect_StringsTranslations_Consolidated.xlsx"))

  // Loop through each sheet
  for (sheetIndex in 0 until workbook.numberOfSheets) {
    val sheet = workbook.getSheetAt(sheetIndex)
    val sheetName = sheet.sheetName

    // Get the column headers
    val headerRow = sheet.getRow(0)
    val columnHeaders = headerRow.cellIterator().asSequence().map { it.stringCellValue }.toList()

    // Find the index of the ID column
    val idColumnIndex = columnHeaders.indexOf("ID")

    // Loop through each language column
    for (languageColumnIndex in columnHeaders.indices) {
      if (languageColumnIndex != idColumnIndex) {
        val languageColumnName = columnHeaders[languageColumnIndex]

        // Create the XML document
        val root = Element("resources")
        val doc = Document(root)

        // Loop through each row and add the string elements
        for (rowIndex in 1 until sheet.lastRowNum + 1) {
          val row = sheet.getRow(rowIndex)
          if (row?.getCell(idColumnIndex) == null) continue
          val id = row.getCell(idColumnIndex).stringCellValue
          val cell = row.getCell(languageColumnIndex) ?: continue
          val languageString = when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> cell.numericCellValue.toString()
            else -> ""
          }
          val stringElement = Element("string")
          stringElement.setAttribute("name", id)
          stringElement.text = languageString
          root.addContent(stringElement)
        }

        // Write the XML document to a file
        val resourcesDir = File("src/main/resources/output/")
        if (!resourcesDir.exists()) resourcesDir.mkdirs()
        val safeSheetName = sheetName.replace(" ", "_").replace("/", "_")
        val safeColumnName = languageColumnName.replace(" ", "_").replace("/", "_")
        val fileName = "strings_${safeSheetName}_$safeColumnName.xml"
        val xmlOutputter = XMLOutputter(Format.getPrettyFormat())
        val fileOutputStream = FileOutputStream(File(resourcesDir, fileName))
        xmlOutputter.output(doc, fileOutputStream)
        fileOutputStream.close()
      }
    }
  }
}