

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
    SPALTEN_UEBERSCHRIFTEN.each do |key, value|
      aktuelle_zelle = @xlapp.WorkSheets(@tabelle_name).Cells(1,1).Find(value[0..6]).Activate
      if aktuelle_zelle
        aktuelle_spalte, aktuelle_zeile = @xlapp.WorkSheets(@tabelle_name).Cells(1,1).Find(value[0..6]).Address.scan(/\w+/)
        erg[key.to_sym] = @xlapp.WorkSheets(@tabelle_name).range("#{aktuelle_spalte}#{zeilen_nummer}").value
      end
    end
    return erg
  end

  def excel_beenden
    @excel_controller.quit_excel
  end
end

