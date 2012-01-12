package ss

import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
//import org.jopendocument.dom.spreadsheet.*
//import org.jopendocument.dom.ODPackage
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.POIXMLDocument

class SpreadsheetBuilderService {

	boolean transactional = true

	def buildXlsFromArray(data) {
		def baos = new ByteArrayOutputStream()
		Workbook wb = new HSSFWorkbook()
		def sheet = wb.createSheet()
		def (row, cell) = [null,null]
		data.eachWithIndex{ rowName, rowData, rowNum ->
			row = sheet.createRow(rowNum)
			println rowName
			println rowData.getClass()
			rowData.eachWithIndex{ cellName, cellData, colNum ->
				cell = row.createCell(colNum)
				cell.setCellValue(cellData)
			}
		}
		wb.write(baos)
		return baos.toByteArray()
	}
	
	def fillInXlsxTemplateWithDataFromArray(templateUrl, data){
		def baos = new ByteArrayOutputStream()
		Workbook wb = new XSSFWorkbook(templateUrl)
		def sheet = wb.getSheetAt(0)
		def (row, cell) = [null,null]
		data.eachWithIndex{ rowName, rowData, rowNum ->
			//println "rowName: ${rowName}"
			//println "rowData: ${rowData}"
			//println "rowNum: ${rowNum}"
			row = sheet.getRow(rowName-1) ?: sheet.createRow(rowName-1)
			rowData.eachWithIndex{ cellName, cellData, colNum ->
				cell = row.getCell(colNum) ?: row.createCell(colNum)
				cell.setCellValue(cellData)
			}
		}
		wb.write(baos)
		return baos.toByteArray()
	}
}
