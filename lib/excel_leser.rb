

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
    #dateiname = FILE_PATH + @excel_datei_name
    dateiname = pfad
    if File.exist?(dateiname) #(@excel_datei_name)
      #      @excel_appl = WIN32OLE.GetActiveObject('Excel.Application') ||
      #        WIN32OLE.new('Excel.Application')
      @excel_appl = WIN32OLE.new('Excel.Application') 
      @excel_appl['Visible'] = true      
      @mappe = @excel_appl.Workbooks.Open(dateiname)
      #@excel_appl.Workbooks.Open(dateiname)
      #@mappe = @excel_appl.ActiveWorkbook #problem wenn alter excel task noch aktiv
      #@mappe = WIN32OLE.connect('Excel.Application').ActiveWorkbook #problem wenn alter excel task noch aktiv
    else
      raise "Error Message: " "Datei '#{dateiname}' nicht vorhanden " # eventuell neue abfrage
    end
  end

  def quit_excel
    @excel_appl.quit
    @excel_appl = nil
    GC.start #garbage control starten
  end
end

class ExcelLeser #< ExcelController
  def initialize(pfad, global_name, tabelle_name)
    @tabelle_name = tabelle_name
    @excel_controller = ExcelController.new(pfad)
    @excel_controller.open_excel_file(pfad)
    @xlapp = @excel_controller.excel_appl
  end
  
  def zeile(zeilen_nummer)
    erg = {}
    erg[:name] = @xlapp.WorkSheets(@tabelle_name).range("A#{zeilen_nummer}").value
    erg[:verzicht_betrag] = @xlapp.WorkSheets(@tabelle_name).range("L#{zeilen_nummer}").value
    erg[:berufsgruppe] = @xlapp.WorkSheets(@tabelle_name).range("Q#{zeilen_nummer}").value
    return erg
  end

  def excel_beenden
    @excel_controller.quit_excel
  end
end

