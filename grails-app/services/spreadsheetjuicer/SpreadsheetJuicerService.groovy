package spreadsheetjuicer

import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.jopendocument.dom.spreadsheet.*
import org.apache.poi.xssf.usermodel.*

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
  
  def getSheetAsHash(uploadFile, columTypes, ignoredRows) {
    def ssData
    def cleanSsData
    switch(uploadFile.contentType) {
      case "application/vnd.ms-excel":
        ssData = xlsToHash(uploadFile)
        cleanSsData = xlsDataCleaner(ssData, columTypes, ignoredRows)
        break
      case "application/vnd.oasis.opendocument.spreadsheet":
        ssData = xlsxToHash(uploadFile)
        cleanSsData = xlsxDataCleaner(ssData, columTypes, ignoredRows)
        break
      case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        ssData = odsToHash(uploadFile)
        cleanSsData = odsDataCleaner(ssData, columTypes, ignoredRows)
        break
      default:
        println "not valid file type?"
    }
    return cleanSsData
  }
  
  def xlsToHash(uploadFile) {
    Workbook wb = new HSSFWorkbook(uploadFile.inputStream)
    //so this is just going to get sheet 1 .. who uses more than 1?
    //def numOfSheets = wb.getNumberOfSheets() 
    Sheet sheet = wb.getSheetAt(0)
    def xlsData = [:]
    sheet.iterator().eachWithIndex{row, rowNum ->
      xlsData[rowNum] = [:]
      row.iterator().eachWithIndex{colum, colNum ->
        xlsData[rowNum][colNum] = colum
      }
    }
    return xlsData
  }
  
  def xlsxToHash(uploadFile) {
    //def wb = new XSSFWorkbook(uploadFile.inputStream)
  }
  
  def xlsxDataCleaner(xlsxData, columTypes, ignoredRows) {
  
  }
  
  def odsToHash(uploadFile) {
    SpreadSheet ss = SpreadSheet.createFromFile(uploadFile.inputStream)
    //def sheetCount = ss.getSheetCount() 
    Sheet sheet = ss.getSheet(0)
    def rowCount = sheet.getRowCount()
    def columnCount = sheet.getColumnCount()
    def odsData = [:]
    (0..getRowCount).each{rowNum ->
      def odsData[rowNum] = [:]
      (0..columnCount).each{colNum ->
        odsData[rowNum][colNum] = sheet.getCellAt(rowNum, colNum)
      }
    }
    return odsData
  }
  
  def odsDataCleaner(odsData, columTypes, ignoredRows) {
  
  }
  
  def xlsDataCleaner(xlsData, columTypes, ignoredRows) {
    def cleanXlsData = [:]
    xlsData.each{rowNum,row ->
      //try { }
      if(!ignoredRows.contains(rowNum)) {
        cleanXlsData[rowNum] = [:]
        row.each{colNum,col ->
          switch ( columTypes[colNum] ) {
            //case ["number","num","int","integer",0, 'inList']
            case "number":
              cleanXlsData[rowNum][colNum] = col.getNumericCellValue()
              println "number: " + col
              break
            case "string":
              cleanXlsData[rowNum][colNum] = col.getStringCellValue()
              println "string: " + col 
              break
            case "date":
              cleanXlsData[rowNum][colNum] = col.getDateCellValue()
              println "date: " + col
              break
            default:
              cleanXlsData[rowNum][colNum] = ''
          }
        }
      }
      else println "ignoring row ${rowNum}"
      //catch { }
    }
    return cleanXlsData
  }
}
