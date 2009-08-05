

require 'win32ole'

require 'grundlage'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class ExcelController
  attr_reader :excel_appl

  def initialize(pfad)
    @excel_datei_name = ""
    @excel_blatt_name = ""
    @excel_appl = ""
    @mappe = ""
  end

  def open_excel_file(pfad)
    dateiname = pfad
    if File.exist?(dateiname)
      #      @excel_appl = WIN32OLE.GetActiveObject('Excel.Application') ||
      #        WIN32OLE.new('Excel.Application')
      @excel_appl = WIN32OLE.new('Excel.Application')
      @excel_appl['Visible'] = true
      @mappe = @excel_appl.Workbooks.Open(dateiname)
    else
      raise "Error Message: " "Datei '#{dateiname}' nicht vorhanden "
      ## eventuell neue abfrage
    end
  end

  def quit_excel
    @excel_appl.quit
    @excel_appl = nil
    GC.start #garbage control starten
  end
end