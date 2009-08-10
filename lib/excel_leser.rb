

require 'win32ole'

require 'grundlage'
require 'excel_controller'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class ExcelLeser #< ExcelController
  def initialize(pfad, global_name, tabelle_name)
    @tabelle_name = tabelle_name
    @global_name = global_name
    WIN32OLE.codepage = WIN32OLE::CP_UTF8 #zeichen als unicode verarbeiten
    @excel_controller = ExcelController.new(pfad)
    @excel_controller.open_excel_file(pfad)
    @xlapp = @excel_controller.excel_appl
    @leerzeilen = 1 # anzahl der leerzeilen zwischen ueberschrift und beginn
    # der datensaetze im excelsheet "Tabelle"
  end

  def zeile_als_array(zeilen_nr) #liest eine zeile aus dem blatt "tabelle" ein
    @xlapp.WorkSheets(@tabelle_name).Rows(zeilen_nr).Value.first
  end

  def namenfeld_wert(namenfeld_bez)
    begin
      erg = @xlapp.WorkSheets(@global_name).Range(namenfeld_bez).Value
    rescue WIN32OLERuntimeError
      nil
    end
  end

  def zeile(zeilen_nummer)
    erg = {}
    blatt_ueberschriften = zeile_als_array(19)
    aktuelle_zeile = zeile_als_array(zeilen_nummer)
    SPALTEN_UEBERSCHRIFTEN.each do |ueberschrift_bezeichnung, ueberschrift_vorgegeben|
      catch :ueberspringen do
        spalten_nr = blatt_ueberschriften.each_with_index do |ueberschrift_aus_blatt, index|
          break nil if ueberschrift_aus_blatt.nil?
          throw :ueberspringen if ueberschrift_vorgegeben.nil?
          regex_neu = Regexp.new(ueberschrift_vorgegeben.source, Regexp::MULTILINE | Regexp::IGNORECASE )
          break index if ueberschrift_aus_blatt.gsub(/[)(]/,"") =~ regex_neu
        end
        if spalten_nr.is_a? Integer
          erg[ueberschrift_bezeichnung] = aktuelle_zeile[spalten_nr]
        else
          raise "Überschrift '#{ueberschrift_vorgegeben}' (für #{ueberschrift_bezeichnung}) nicht gefunden."
        end
      end
    end
    GLOBALBLATT_NAMEN.each do |namenfeld_bezeichnung, namenfeld_vorgegeben|
      aktuelles_namenfeld = namenfeld_wert(namenfeld_vorgegeben)
      erg[namenfeld_bezeichnung] = aktuelles_namenfeld
    end
    EXCEL_EINLESE_TRANSFORMATIONEN.each do |key, trafo_hash|
      alter_wert = erg[key]
      neuer_wert = trafo_hash[alter_wert]
      erg[key] = neuer_wert
     #alternative: erg[key] = trafo_hash[erg[key]]
    end
    return erg
  end

  def excel_beenden
    @excel_controller.quit_excel
  end
end

