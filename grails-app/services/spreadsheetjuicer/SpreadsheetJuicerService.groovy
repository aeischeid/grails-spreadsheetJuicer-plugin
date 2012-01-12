package ss

import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.jopendocument.dom.spreadsheet.*
import org.jopendocument.dom.ODPackage
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.POIXMLDocument

class SpreadsheetJuicerService {

	boolean transactional = true
	
	def getFileProperties(uploadFile) {
		def fileProps = [:] 
		//log.debug "content type:	" + uploadFile.contentType
		switch(uploadFile.contentType) {
			case "application/vnd.ms-excel":
				fileProps.type = "xls"
				break
			case "application/vnd.oasis.opendocument.spreadsheet":
				fileProps.type = "ods"
				break
			case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
				fileProps.type = "xlsx"
				break
			default:
				fileProps.type = ""
		}
		fileProps.size = uploadFile.size
//		log.debug fileProps
		return fileProps
	}
	
	def getSheetAsSimpleHash(uploadFile, columTypes, ignoredRows) {
		def cleanSsData
		switch(uploadFile.contentType) {
			case "application/vnd.ms-excel":
				cleanSsData = xlsToSimpleHash(uploadFile.inputStream, columTypes, ignoredRows)
				break
			case "application/vnd.oasis.opendocument.spreadsheet":
				cleanSsData = odsToSimpleHash(uploadFile.inputStream, columTypes, ignoredRows)
				break
			case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
				cleanSsData = xlsxToSimpleHash(uploadFile.inputStream, columTypes, ignoredRows)
				break
			default:
				log.debug "not valid file type?"
		}
		return cleanSsData
	}

	def getSheetAsHash(uploadFile, columTypes, ignoredRows) {
		def cleanSsData
		switch(uploadFile.contentType) {
			case "application/vnd.ms-excel":
				cleanSsData = excelToHash(uploadFile.inputStream, columTypes, ignoredRows)
				break
			case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
				cleanSsData = excelToHash(uploadFile.inputStream, columTypes, ignoredRows)
				break
			case "application/vnd.oasis.opendocument.spreadsheet":
				cleanSsData = odsToHash(uploadFile.inputStream, columTypes, ignoredRows)
				break
			default:
				log.debug "not valid file type?"
		}
		return cleanSsData
	}
	
//		Idealy we want the resulting hash to be in this format
//			var documents = [
//				{ //repeats
//						metadata: {
//								columns: Column_Count,
//								rows: Row_Count,
//								title: ''
//						},
//						data: {
//								r{Row_Index}: { //repeats
//										c{Column_Index}: { //repeats
//												value: '',
//												style: '',
//												width: 0,
//												cl: {Classes used for styling}
//										}
//								}
//						}
//				}
//			]
	
	def xlsToSimpleHash(inputStream, columTypes, ignoredRows) {
		Workbook wb = new HSSFWorkbook(inputStream)
		//so this is just going to get sheet 1 .. who uses more than 1?
		//def numOfSheets = wb.getNumberOfSheets() 
		Sheet sheet = wb.getSheetAt(0)
		def xlsData = [:]
		sheet.iterator().eachWithIndex{row, rowNum ->
			if(!ignoredRows.contains(rowNum)) {
				xlsData[rowNum] = [:]
				row.iterator().eachWithIndex{col, colNum ->
					switch ( columTypes[colNum] ) {
						//case ["number","num","int","integer",0, 'inList']
						case "number":
							xlsData[rowNum][colNum] = col.getNumericCellValue()
							break
						case "string":
							xlsData[rowNum][colNum] = col.getStringCellValue()
							break
						case "date":
							xlsData[rowNum][colNum] = col.getDateCellValue()
							break
						default:
							xlsData[rowNum][colNum] = ''
					}
					//log.debug xlsData[rowNum][colNum] + " - class: "xlsData[rowNum][colNum].class
				}
			}
			else log.debug "ignoring row ${rowNum}"
		}
		return xlsData
	}
	
	def xlsxToSimpleHash(inputStream, columTypes, ignoredRows) {
		//def wb = new XSSFWorkbook(inputStream)
	}
	
	def odsToSimpleHash(inputStream, columTypes, ignoredRows) {
		ODPackage odp = new ODPackage(inputStream)
		SpreadSheet ods = SpreadSheet.create(odp)
		//def sheetCount = ss.getSheetCount() 
		def sheet = ods.getSheet(0)
		// -1 because we want resulting hash to start counting at 0 for consistancy with the way we handle xls files
		def rowCount = sheet.getRowCount()-1
		def columnCount = columTypes.size()-1 
		//log.debug "rows: "+ rowCount + " -- columns: " + columnCount
		def odsData = [:]
		(0..rowCount).each{rowNum ->
			if(!ignoredRows.contains(rowNum)) {
				odsData[rowNum] = [:]
				(0..columnCount).each{colNum ->
					//getValue returns a java object! so simple!
					odsData[rowNum][colNum] = sheet.getCellAt(colNum, rowNum).getValue()
				}
			}
			else log.debug "ignoring row ${rowNum}"
		}
		return odsData
	}
	
	def xlsToHash(inputStream, colTypes, ignoredRows) {
		Workbook wb = new HSSFWorkbook(inputStream)
		//so this is just going to get sheet 1 .. who uses more than 1?
		//def numOfSheets = wb.getNumberOfSheets() 
		def xlsData = []
		Sheet sheet = wb.getSheetAt(0)
		def rowCount = sheet.getLastRowNum()+1
		def colCount = sheet.getRow(0).getLastCellNum()

		xlsData[0] = ["metadata":["columns":colCount, "rows":rowCount, "title":sheet.getSheetName()], "data":[:]]
		sheet.iterator().eachWithIndex{row, rowNum ->
			xlsData[0]["data"]['r'+(rowNum+1)] = [:]
			row.iterator().eachWithIndex{cell, colNum ->
				// get cellType returns 1 for string 0 for numeric (and 2 for formula?)
				if(cell.getCellType()){
					try{xlsData[0]["data"]['r'+(rowNum+1)]['c'+(colNum+1)] = ["value":cell.getStringCellValue(),"style":"","width":0,"cl":[]]}
					catch(Throwable e){log.debug e}
				}
				else{
					// excel stores dates as a formatted Double. check the format to see if it is a date cell
					if( new DateUtil().isCellDateFormatted(cell) ){
						try{xlsData[0]["data"]['r'+(rowNum+1)]['c'+(colNum+1)] = ["value":cell.getDateCellValue(),"style":"","width":0,"cl":[]]}
						catch(Throwable e){log.debug e}
					}
					else {
						try{xlsData[0]["data"]['r'+(rowNum+1)]['c'+(colNum+1)] = ["value":cell.getNumericCellValue(),"style":"","width":0,"cl":[]]}
						catch(Throwable e){log.debug e}
					}
				} 
			}
		}
		return xlsData
	}
	
	def excelToHash(inputStream, rows = null, columns = null, parseDatesAsString = false) {
		//log.debug "SpreadsheetJuicerService ****************************************************"
		//log.debug "    excelToHash"
		def wb = new WorkbookFactory().create(inputStream)
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator()
		def (dateUtil, excelData, sheet, row, cell, cellVal) = [new DateUtil(), [], null,null,null,null]
		def numOfSheets = wb.getNumberOfSheets()-1
		(0..numOfSheets).each{sheetNum ->
			sheet = wb.getSheetAt(sheetNum)
			def firstRow = sheet.getFirstRowNum()
			if(sheet.getRow(firstRow)){
				// default tries to auto determine last row... option to specify rows
				// not sure what it will do with blank rows acting as spacers
				def rowCount = rows ?: sheet.getLastRowNum()+1
				// default grid width is determined by the first row
				// this is not always a good situation... option to specify columns
				def colCount = columns ?: sheet.getRow(firstRow).getLastCellNum()
				//log.debug "   excelToHash:  Rows = " + rowCount + ", Columns = " + colCount
				excelData[sheetNum] = [
					"metadata":["columns":colCount, "rows":rowCount,"title":sheet.getSheetName()],
					"data":[:]
				]
				// should we start building array from row0 even if it is technically blank?
				// if we don't jquery sheet plugin may have problems...
				//for(rowNum in firstRow..rowCount) {
				for(rowNum in 0..rowCount) {
					excelData[sheetNum]["data"]['r'+(rowNum)] = [:]
					row = sheet.getRow(rowNum)
					if(row){
						// should we even bother with determining firstCell? Starting at 0 is easier and more consistent
						//def firstCell = row.getFirstCellNum() ?: 0
						//for(colNum in firstCell..colCount){
						for(colNum in 0..colCount){
							cell = row.getCell(colNum, Row.CREATE_NULL_AS_BLANK)
							def kindOfCell = cell.getCellType() // 0 = numeric, 1 = String, 2 = Formula, 3 = Blank, 4 = Boolean, 5 = Error
							switch(kindOfCell) {
								case 0:
									//log.debug "numeric cellFormat: "+cell.getCellStyle().getDataFormat() + " --- " +cell.getCellStyle().getDataFormatString()
									if(!dateUtil.isCellDateFormatted(cell)) cellVal = cell.getNumericCellValue()
									else {
										if(!parseDatesAsString) cellVal = dateUtil.getJavaDate(cell.getNumericCellValue())
										else cellVal = new DataFormatter().formatCellValue(cell)
									}
									break
								case 1:
									//log.debug "string cellFormat: "+cell.getCellStyle().getDataFormat() + " --- " +cell.getCellStyle().getDataFormatString()
									cellVal = cell.getStringCellValue()
									break
								case 2:
									//cellVal = cell.getCachedFormulaResultType().toString()
									def cellValue = evaluator.evaluate(cell)
									switch (cellValue.getCellType()) {
										case Cell.CELL_TYPE_BOOLEAN:
												cellVal = cellValue.getBooleanValue()
												break;
										case Cell.CELL_TYPE_NUMERIC:
												def cellNumVal = cellValue.getNumberValue()
												if(!dateUtil.isCellDateFormatted(cell)) cellVal = cellNumVal
												else {
													if(!parseDatesAsString) cellVal = dateUtil.getJavaDate(cellNumVal)
													else {
														try{
															def dd = cellNumVal.doubleValue()
															def ii = cell.getCellStyle().getIndex().intValue()
															cellVal = new DataFormatter().formatRawCellcontents(dd,ii, cell.getCellStyle().getDataFormatString())
														}
														catch(e){ cellVal = cellNumVal }
														
													}
												}
												break;
										case Cell.CELL_TYPE_STRING:
												cellVal = cellValue.getStringValue()
												break;
										case Cell.CELL_TYPE_BLANK:
												cellVal = ''
												break;
										case Cell.CELL_TYPE_ERROR:
												cellVal = 'err'
												break;
									}
									break
								case 3:
									cellVal = ''
									break
								case 4:
									cellVal = cell.getBooleanCellValue()
									break
								default:
									cellVal = "error"
									log.debug "cell evaluation error"
							}
							excelData[sheetNum]["data"]['r'+(rowNum)]['c'+(colNum)] = ["value":cellVal,
																																				 "style":"",
																																				 "width":0,
																																				 "cl":["cell"]]
						}
					}
					else {
						for(colNum in 0..colCount){
							excelData[sheetNum]["data"]['r'+(rowNum)]['c'+(colNum)] = ["value":'',
																																				 "style":"",
																																				 "width":0,
																																				 "cl":["cell"]]
						}
					}
				}
			}
		}
		return excelData
	}
	
	def odsToHash(inputStream, colTypes, ignoredRows) {
		ODPackage odp = new ODPackage(inputStream)
		SpreadSheet ods = SpreadSheet.create(odp)
		//def sheetCount = ss.getSheetCount() 
		def sheet = ods.getSheet(0)
		//log.debug "style name: "+ sheet.getStyleName()
		def rowCount = sheet.getRowCount()
		// having a tough time with colCount
		//log.debug "colCount = "+ sheet.getColumnCount()
		def colCount = colTypes.size()
		//log.debug "rows: "+ rowCount + " -- columns: " + colCount
		def odsData = []
		odsData[0] = ["metadata":["columns":colCount, "rows":rowCount, "title":1],
"data":[:]]
		(1..rowCount).each{rowNum ->
			odsData[0]["data"]['r'+rowNum] = [:]
			(1..colCount).each{colNum ->
				//getValue returns a java object! so simple!
				//log.debug sheet.getCellAt(colNum-1, rowNum-1).getValue()
				odsData[0]["data"]['r'+rowNum]['c'+colNum] = sheet.getCellAt(colNum-1, rowNum-1).getValue()
			}
		}
		return odsData
	}
}
