Grails SpreadsheetJuicer Plugin
author: Aaron Eischeid

The spreadsheetJuicer plugin provides Grails with a service to extract the
data from simple MSexcel and OpenOffice spreadsheets into a form easily used
in groovy. Under the hood it is using apache POI (3.7) and JOpenDocument (1.2) 
libraries.

Basically it is a convience plugin to save people with relatively simple 
spreadsheets from having to work with the underlying libs which are very 
complete and can thus be a bit complex.

Additionally I have worked a bit with the developer of jQuery.sheet, a plugin for 
displaying spreadsheet data in a browser. http://www.visop-dev.com/jquerysheet.html
The resulting data format is very simple to convert to JSON or XML and import into 
that plugin for manipulation in browser

<------------>
    Usage:
<------------>

Juicer:

def uploadedFile = new File("${params.fileName}")
def fileStream = new FileInputStream(uploadedFile)
return spreadsheetJuicerService.excelToHash(fileStream)

Builder:

// build an array of data you want to fill the spreadsheet with
def books = Books.list
def ssData = [:]
books.eachWithIndex{
	ssData[i] = [a:"${it.author.name}",
							b:"${it.title}",
							c:"${it.pagecount}",
							.
							.
							.
							t:'${it.publisher.name}']
}
// get path to a premade excel sheet that we will populate
// row 1 is assumed to be colum headings so service starts filling in on row 2
def b = spreadsheetBuilderService.fillInXlsxTemplateWithDataFromArray(templateUrl, ssData)
response.contentType = ConfigurationHolder.config.grails.mime.types['excel']
response.setHeader("Content-disposition", "attachment; filename=billing-week${params.weekOfYear}-${params.year}.xlsx")
response.setContentLength(b.length)
response.getOutputStream().write(b)
