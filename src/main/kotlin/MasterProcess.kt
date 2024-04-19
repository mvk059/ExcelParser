package org.mvk

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.jdom2.Element
import org.jdom2.output.Format
import org.jdom2.output.XMLOutputter
import java.io.File
import java.io.FileOutputStream
import java.util.regex.Pattern

fun main() {
  val matchThreshold = 0.8 // Match threshold, configurable

  val resourceFolder = "src/main/resources"
  // Read the Excel file
  val excelFile = File("$resourceFolder/copy.xlsx")
  val workbook = WorkbookFactory.create(excelFile)

  // Read the original XML file
  val originalXmlFile = File("$resourceFolder/strings.xml")
  val doc = org.jdom2.input.SAXBuilder().build(originalXmlFile)
  val root = doc.rootElement
  var size = 0

  // Get the unique language names from all sheets
  val uniqueLanguages = workbook.sheetIterator().asSequence()
    .flatMap { sheet ->
      sheet.getRow(0).cellIterator().asSequence().drop(1).map { getCellValue(it) }
    }
    .distinct()
    .toList()

  // Create output XML elements for each unique language
  val outputXmls = uniqueLanguages.associateWith { Element("resources") }

  root.children.forEach { stringElement ->
    val id = stringElement.getAttributeValue("name")
    val englishString = stringElement.text.trim().removeSubscriptSuperscript().removeFormatSpecifiers()
    val isTranslatable = stringElement.getAttributeValue("translatable")?.toBoolean() ?: true

    if (isTranslatable) {
      val matchingRow = workbook.sheetIterator().asSequence()
        .flatMap { sheet ->
          val sheetLanguages = sheet.getRow(0).cellIterator().asSequence().drop(1).map { getCellValue(it) }.toList()
          sheet.rowIterator().asSequence()
            .map { row ->
              row.cellIterator().asSequence()
                .map { getCellValue(it).removeSubscriptSuperscript().removeFormatSpecifiers() }.toList()
            }
            .filter { it.size == sheetLanguages.size + 1 }
        }
        .firstOrNull { englishString.similarityScore(it.first()) >= matchThreshold }

      if (matchingRow != null) {
        uniqueLanguages.forEachIndexed { languageIndex, language ->
          val columnIndex = uniqueLanguages.indexOfFirst { it == language }
          if (columnIndex != -1) {
            val translatedString = matchingRow[columnIndex + 1] // +1 to skip the English column
            val childElement = Element("string")
            childElement.setAttribute("name", id)
            childElement.text = translatedString
            outputXmls[language]?.addContent(childElement)
          }
        }
      } else {
        size = size.inc()
        println("$size: $id - $englishString")
      }
    }
  }

  // Write the output XML files
  outputXmls.forEach { (language, outputXml) ->
    val outputFile = File("$resourceFolder/output2/$language.xml")
    outputFile.parentFile.mkdirs()
    val xmlOutputter = XMLOutputter(Format.getPrettyFormat())
    xmlOutputter.output(outputXml, FileOutputStream(outputFile))
  }
  println("Size: $size")

}

fun getCellValue(cell: org.apache.poi.ss.usermodel.Cell): String {
  return when (cell.cellType) {
    CellType.STRING -> cell.stringCellValue
    CellType.NUMERIC -> cell.numericCellValue.toString()
    else -> ""
  }
}

fun String.similarityScore(other: String): Double {
  val longerLength = maxOf(this.length, other.length)
  if (longerLength == 0) return 1.0
  return (2.0 * longerLength(this, other)) / (this.length + other.length)
}

fun String.removeSubscriptSuperscript(): String {
  val pattern = Pattern.compile("[\\u2070-\\u209F\\u2080-\\u209F-\\u0032]")
  return pattern.matcher(this).replaceAll("")
}

fun String.removeFormatSpecifiers(): String {
  val pattern = Pattern.compile("%\\d+\\$\\w")
  return pattern.matcher(this).replaceAll("")
}

fun longerLength(a: String, b: String): Int {
  val dp = Array(a.length + 1) { IntArray(b.length + 1) }

  for (i in 0 until a.length + 1) {
    for (j in 0 until b.length + 1) {
      if (i == 0 || j == 0) {
        dp[i][j] = 0
      } else if (a[i - 1] == b[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1
      } else {
        dp[i][j] = maxOf(dp[i - 1][j], dp[i][j - 1])
      }
    }
  }

  return dp[a.length][b.length]
}