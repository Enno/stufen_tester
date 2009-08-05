

require 'win32ole'

require 'grundlage'
require 'excel_controller'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class FormFiller
  attr_reader :excel_appl

  def initialize(dateiname, start_proc_name)
    WIN32OLE.codepage = WIN32OLE::CP_UTF8 #zeichen als unicode verarbeiten
    @excel_controller = ExcelController.new(dateiname)
    @excel_controller.open_excel_file(dateiname)
    @xlapp = @excel_controller.excel_appl
  end

  def fill(zeile)

  end

  def ergebnis_anfordern
    @excel_controller.quit_excel
  end

end

