

require 'win32ole'

require 'grundlage'
require 'excel_controller'
require 'tasten_sender'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class FormFiller 
 
  def initialize(pfad, dateiname, start_proc_name)
    @proc_name = start_proc_name
    @datei_name = dateiname
    @excel_controller = ExcelController.new(pfad + dateiname)
    @excel_controller.open_excel_file(pfad + dateiname)
    @xlapp = @excel_controller.excel_appl
    @masken_controller = TastenSender.new()
    @feld_kinderlos_aktiv = true
  end

  def maske_oeffnen
    #@xlapp.Run "#{@datei_name}!#{@proc_name}"
    @masken_controller.sende_tasten('Microsoft Excel', "%{F8}#{@proc_name}%{a}")
  end

  def naechstes_feld(anzahl)
    anzahl_tabs = "{TAB}"
    anzahl_tabs = anzahl_tabs*anzahl
    TastenSender.new().sende_tasten('Microsoft Excel', "#{anzahl_tabs}")
  end
  def eingabe_bestaetigen
    TastenSender.new().sende_tasten('Microsoft Excel', "{ENTER}", :wartezeit => 1)
  end

  def maske_fuellen(zeile)
    maske_oeffnen
    @masken_fueller = TastenSender.new()
    #Blatt 1
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:name]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:bruttogehalt]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:freibetrag]}")
    naechstes_feld(1)
    if zeile[:k_vers_art] == "p"
      naechstes_feld(1)
      @masken_fueller.sende_tasten('Microsoft Excel', " ")
      feld_kinderlos_aktiv = false
    else
      @masken_fueller.sende_tasten('Microsoft Excel', " ")
      naechstes_feld(1)
      feld_kinderlos_aktiv = true
    end 
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:steuerklasse]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:kinder_fb]}")
    naechstes_feld(1)
    if zeile[:kirchensteuer] == "n"
      @masken_fueller.sende_tasten('Microsoft Excel', " ")
    end
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:bland_wohnsitz]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:bland_arbeit]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:berufsgruppe]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:durchfuehrungsweg]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:pausch_steuer40b]}")
    naechstes_feld(1)
    @masken_fueller.
      sende_tasten('Microsoft Excel', "#{zeile[:minijob_ok ]}")
    naechstes_feld(1)
    if feld_kinderlos_aktiv == true
      if zeile[:kinderlos] == "j"
        @masken_fueller.sende_tasten('Microsoft Excel', " ")
      end
      naechstes_feld(1)
    end
    eingabe_bestaetigen
    eingabe_bestaetigen
    naechstes_feld(1)
    # @masken_fueller.sende_tasten('Microsoft Excel', "%j", :wartezeit => 1)
    #Blatt 2
  end

  def vb_senden(vb_abfrage)
    @xlapp.Run "#{@datei_name}!#{vb_abfrage}" #(vb_abfrage)
  end

  def maske_schliessen
    @xlapp.ActiveWorkbook.Close
  end

  def excel_beenden
    @excel_controller.quit_excel
  end

end

