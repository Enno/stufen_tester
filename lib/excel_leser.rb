

require 'win32ole'

require 'grundlage'
require 'excel_controller'
require 'tab_zeile'

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

  def werte_auf_integer_pruefen(aktuelle_zeile)
    aktuelle_zeile.each_with_index do |wert, index|
      aktuelle_zeile[index] = wert.to_i if wert.is_a?(Float) && wert.to_s.match(/[.0]+$/)
    end
    return aktuelle_zeile
  end

  def zeile(zeilen_nummer)
    @erg = TabZeile.new
    
    aktuelle_zeile = zeile_als_array(zeilen_nummer)
    werte_auf_integer_pruefen(aktuelle_zeile)
    lese_pers_daten(aktuelle_zeile)
    GLOBALBLATT_NAMEN.each do |namenfeld_bezeichnung, namenfeld_vorgegeben|
      aktuelles_namenfeld = namenfeld_wert(namenfeld_vorgegeben)
      @erg.eingaben[namenfeld_bezeichnung] = aktuelles_namenfeld
    end
    EXCEL_EINLESE_TRANSFORMATIONEN.each do |key, trafo_hash|
      alter_wert = @erg.eingaben[key]
      neuer_wert = trafo_hash[alter_wert]
      @erg.eingaben[key] = neuer_wert
      #alternative: erg[key] = trafo_hash[erg[key]]
    end
    return @erg
  end

  def lese_pers_daten(aktuelle_zeile)
    bereich_ueberschriften = zeile_als_array(17)
    spalten_ueberschriften = zeile_als_array(19)
    BEREICHE_INTERN_ZU_EXCEL.each do |int_bereich, excel_bereich|
      SPALTEN_UEBERSCHRIFTEN_TEST[int_bereich].each do |ueberschrift_bezeichnung, ueberschrift_vorgegeben|
        #p [excel_bereich, ueberschrift_bezeichnung, ueberschrift_vorgegeben]
        if ueberschrift_vorgegeben
#        catch :ueberspringen do
          spalten_nr = spalten_ueberschriften.each_with_index do |ueberschrift_aus_blatt, akt_spalten_nr|
            #p [akt_spalten_nr, ueberschrift_aus_blatt]
            next nil if ueberschrift_aus_blatt.nil?
 #           throw :ueberspringen if (ueberschrift_vorgegeben.nil?)
            regex_neu = Regexp.new(ueberschrift_vorgegeben.source, Regexp::MULTILINE | Regexp::IGNORECASE )
            if bereich_ueberschriften[akt_spalten_nr] == excel_bereich and
               ueberschrift_aus_blatt.gsub(/[)(]/,"") =~ regex_neu
              break akt_spalten_nr
            end
          end
          if spalten_nr.is_a? Integer
            @erg[ueberschrift_bezeichnung, excel_bereich] = aktuelle_zeile[spalten_nr]
          else
            raise "Überschrift '#{ueberschrift_vorgegeben}' (für #{ueberschrift_bezeichnung}) nicht in #{int_bereich.inspect} gefunden."
          end
        end
      end
    end
  end

  def excel_beenden
    @excel_controller.quit_excel
  end
end

