package com.gvl

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
    //println "content type:  " + uploadFile.contentType
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
    println fileProps
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
        println "not valid file type?"
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
        println "not valid file type?"
    }
    return cleanSsData
  }
  
//    Idealy we want the resulting hash to be in this format -- not necessarily JSON syntax. 
//      var documents = [
//        { //repeats
//            metadata: {
//                columns: Column_Count,
//                rows: Row_Count,
//                title: ''
//            },
//            data: {
//                r{Row_Index}: { //repeats
//                    c{Column_Index}: { //repeats
//                        value: '',
//                        style: '',
//                        width: 0,
//                        cl: {Classes used for styling}
//                    }
//                }
//            }
//        }
//      ]
  
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
          //println xlsData[rowNum][colNum] + " - class: "xlsData[rowNum][colNum].class
        }
      }
      else println "ignoring row ${rowNum}"
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
    println "rows: "+ rowCount + " -- columns: " + columnCount
    def odsData = [:]
    (0..rowCount).each{rowNum ->
      if(!ignoredRows.contains(rowNum)) {
        odsData[rowNum] = [:]
        (0..columnCount).each{colNum ->
          //getValue returns a java object! so simple!
          odsData[rowNum][colNum] = sheet.getCellAt(colNum, rowNum).getValue()
        }
      }
      else println "ignoring row ${rowNum}"
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
    println "rows: "+ rowCount + " -- columns: " + colCount
    xlsData[0] = ["metadata":["columns":colCount, "rows":rowCount, "title":sheet.getSheetName()], "data":[:]]
    sheet.iterator().eachWithIndex{row, rowNum ->
      xlsData[0]["data"]['r'+(rowNum+1)] = [:]
      row.iterator().eachWithIndex{cell, colNum ->
        // get cellType returns 1 for string 0 for numeric (and 2 for formula?)
        if(cell.getCellType()){
          try{xlsData[0]["data"]['r'+(rowNum+1)]['c'+(colNum+1)] = ["value":cell.getStringCellValue(),"style":"","width":0,"cl":[]]}
          catch(Throwable e){println e}
        }
        else{
          // excel stores dates as a formatted Double. check the format to see if it is a date cell
          if( new DateUtil().isCellDateFormatted(cell) ){
            try{xlsData[0]["data"]['r'+(rowNum+1)]['c'+(colNum+1)] = ["value":cell.getDateCellValue(),"style":"","width":0,"cl":[]]}
            catch(Throwable e){println e}
          }
          else {
            try{xlsData[0]["data"]['r'+(rowNum+1)]['c'+(colNum+1)] = ["value":cell.getNumericCellValue(),"style":"","width":0,"cl":[]]}
            catch(Throwable e){println e}
          }
        } 
      }
    }
    return xlsData
  }
  
  def excelToHash(inputStream, colTypes, ignoredRows) {
    def wb = new WorkbookFactory().create(inputStream)
    def excelData = []
    def numOfSheets = wb.getNumberOfSheets()-1
    (0..numOfSheets).each{sheetNum ->
      def sheet = wb.getSheetAt(sheetNum)
      //the sheet needs to have at least one row... not sure what is best detector...
      //maybe if(sheet.getPhysicalNumberOfRows()>0 || sheet.getFirstRowNum())
      //println "physical rows: "+ sheet.getPhysicalNumberOfRows()
      def firstRow = sheet.getFirstRowNum()
      if(sheet.getRow(firstRow)){
        def rowCount = sheet.getLastRowNum()+1
        def colCount = sheet.getRow(firstRow).getLastCellNum()
        //println "rows: "+ rowCount + " -- columns: " + colCount
        excelData[sheetNum] = [
          "metadata":["columns":colCount, "rows":rowCount,"title":sheet.getSheetName()],
          "data":[:]
        ]
        sheet.rowIterator().eachWithIndex{row, rowNum ->
          excelData[sheetNum]["data"]['r'+(rowNum)] = [:]
          row.cellIterator().eachWithIndex{cell, colNum ->
            def kindOfCell = cell.getCellType() // 0 = numeric, 1 = String, 2 = Formula, 3 = Blank, 4 = Boolean, 5 = Error
            def cellVal
            switch(kindOfCell) {
              case 0:
                cellVal = cell.getNumericCellValue()
                break
              case 1:
                cellVal = cell.getStringCellValue()
                break
              case 2:
                cellVal = cell.getCachedFormulaResultType()
                break
              case 3:
                cellVal = ''
                break
              case 4:
                cellVal = cell.getBooleanCellValue()
                break
              default:
                cellVal = "error"
                println "cell evaluation error"
            }
            try{
              excelData[sheetNum]["data"]['r'+(rowNum)]['c'+(colNum)] = ["value":"${cellVal}",
                                                                         "style":"",
                                                                         "width":0,
                                                                         "cl":["cell"]]
            } catch(Throwable e) {
              println 'sheet:'+sheetNum+' -- r:'+rowNum+' -- c:'+colNum+' -- err:'+e
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
    println "style name: "+ sheet.getStyleName()
    def rowCount = sheet.getRowCount()
    // having a tough time with colCount
    println "colCount = "+ sheet.getColumnCount()
    def colCount = colTypes.size()
    println "rows: "+ rowCount + " -- columns: " + colCount
    def odsData = []
    odsData[0] = ["metadata":["columns":colCount, "rows":rowCount, "title":1],
"data":[:]]
    (1..rowCount).each{rowNum ->
      odsData[0]["data"]['r'+rowNum] = [:]
      (1..colCount).each{colNum ->
        //getValue returns a java object! so simple!
        //println sheet.getCellAt(colNum-1, rowNum-1).getValue()
        odsData[0]["data"]['r'+rowNum]['c'+colNum] = sheet.getCellAt(colNum-1, rowNum-1).getValue()
      }
    }
    return odsData
  }
}
