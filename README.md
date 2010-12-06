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

 -- not packaged as a plugin yet, but when it is you will be able to ...
Install from the zip file on Github.
    "grails install-plugin [path/to/zip/file]"
 -- I will also try to release it to the grails repos with relase-plugin --zipOnly ..."

<------------>
    Usage:
<------------>

def webRootDir = servletContext.getRealPath("/")
def uploadedFile = new File("${userDir}/path/to/file")
def file = new FileInputStream(uploadedFile)
def ssDataHash = SpreadsheetJuicerService.xlsAsHash(file)

//simple conversion to JSON 

def ssJson = ssDataHash.encodeAsJSON()
return [params:params, ssJson:ssJson]



a more thorough example:

in gsp --

   
    
in controller -- 

class UploadDataController {

    def spreadsheetJuicerService
    
}
    
back in the gsp --
    
    [all the other pieces you need for the jQuery.sheet plugin]
    

