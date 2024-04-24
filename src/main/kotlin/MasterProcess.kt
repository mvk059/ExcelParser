package org.mvk

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.jdom2.Element
import org.jdom2.output.Format
import org.jdom2.output.XMLOutputter
import java.io.File
import java.io.FileOutputStream
import java.util.regex.Pattern

/*
 * Check each string of each language to verify
 * Now checking for the full string, hence matchThreshold = 1
 * Not removing subscripts and %1$s
 */
fun main() {
    val matchThreshold = 1 // Match threshold, configurable

    val resourceFolder = "src/main/resources"
    // Read the Excel file
    val excelFile = File("$resourceFolder/copy.xlsx")
    val workbook = WorkbookFactory.create(excelFile)

    // Read the original XML file
    val originalXmlFile = File("$resourceFolder/strings.xml")
    val doc = org.jdom2.input.SAXBuilder().build(originalXmlFile)
    val root = doc.rootElement

    val set = mutableSetOf<Pair<String, String>>()

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
        val englishString = stringElement.text.trim()
        val isTranslatable = stringElement.getAttributeValue("translatable")?.toBoolean() ?: true

        if (isTranslatable) {
            val matchingRow = workbook.sheetIterator().asSequence()
                .flatMap { sheet ->
                    val sheetLanguages =
                        sheet.getRow(0).cellIterator().asSequence().drop(1).map { getCellValue(it) }.toList()
                    sheet.rowIterator().asSequence()
                        .map { row ->
                            row.cellIterator().asSequence()
                                .map { getCellValue(it)/*.removeSubscriptSuperscript().removeFormatSpecifiers()*/ }.toList()
                        }
                        .filter { it.size == sheetLanguages.size + 1 }
                }
                .firstOrNull {
                    englishString//.removeSubscriptSuperscript().removeFormatSpecifiers()
                        .similarityScore(it.first()) >= matchThreshold
                }

            uniqueLanguages.forEachIndexed { _, language ->
                val columnIndex = uniqueLanguages.indexOfFirst { it == language }
                if (columnIndex != -1) {
                    val translatedString =
                        if (matchingRow != null) matchingRow[columnIndex + 1].removeExcelWhitespaces() // +1 to skip the English column
                        else {
                            set.add(id to englishString)
                            "$englishString TODO"
                        }
                    val childElement = Element("string")
                    childElement.setAttribute("name", id)
                    childElement.text = translatedString
                    outputXmls[language]?.addContent(childElement)
                }
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
    set.forEach {
        println(it)
    }
    println("Size: ${set.size}")
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
    val pattern = Pattern.compile("%\\d+\\$\\w")    // %1$s, %2$s ...
    return pattern.matcher(this).replaceAll("")
}

fun String.removeExcelWhitespaces(): String {
    return this.replace(" ", " ")
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