

require 'win32ole'

require 'grundlage'
require 'excel_controller'
require 'tasten_sender'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class FormFiller 
 
  def initialize(pfad, dateiname, start_proc_name)
    @dateiname = dateiname
    @proc_name = start_proc_name
    @excel_controller = ExcelController.new(pfad + dateiname)
    @excel_controller.open_excel_file(pfad + dateiname)
    @xlapp = @excel_controller.excel_appl
    @masken_controller = TastenSender.new(:wartezeit => 0.02)

    @feld_kinderlos_aktiv = true
    @feld_pauschalverst_aktiv = true
    @feld_minijob_aktiv = true
    @fenstername = "Microsoft Excel - #{@dateiname}"

    #@masken_fueller = TastenSender.new()
  end

  def maske_oeffnen
    #@xlapp.Run "#{@datei_name}!#{@proc_name}"
    @masken_controller.sende_tasten(@fenstername, "%{F8}#{@proc_name}%{a}", :wartezeit => 0.1, :fenster_fehlt=>"Komischerweise fehlt das Excel-Fenster")
    #sleep(2)
  end

  def feld_vor(anzahl)
    anzahl_tabs = "{TAB}"
    anzahl_tabs = anzahl_tabs*anzahl
    @masken_controller.sende_tasten(@fenstername, "#{anzahl_tabs}")
  end

  def feld_zurueck(anzahl)
    anzahl_tabs = "+{TAB}"
    anzahl_tabs = anzahl_tabs*anzahl
    @masken_controller.sende_tasten(@fenstername, "#{anzahl_tabs}")
  end

  def eingabe_bestaetigen
    @masken_controller.sende_tasten(@fenstername, "{ENTER}") #, :wartezeit => 0.1)
  end

  def zeichen_senden(zeichen)    
    @masken_controller.sende_tasten(@fenstername, "#{zeichen}")
  end

  def maske_fuellen(zeile)
    maske_oeffnen
    #Blatt 1
    zeichen_senden(zeile[:name]);  feld_vor(1)
    zeichen_senden(zeile[:bruttogehalt]);  feld_vor(1)
    zeichen_senden(zeile[:freibetrag]);  feld_vor(1)
    if zeile[:k_vers_art] == "p"
      feld_vor(1);  zeichen_senden(' ')
      @feld_kinderlos_aktiv = false
    else
      zeichen_senden(' ');  feld_vor(1)
    end 
    feld_vor(1)
    zeichen_senden(zeile[:steuerklasse]);  feld_vor(1)
    zeichen_senden(zeile[:kinder_fb]);  feld_vor(1)
    zeichen_senden(' ') if zeile[:kirchensteuer] == "n"
    feld_vor(1)
    zeichen_senden(zeile[:bland_wohnsitz]);  feld_vor(1)
    zeichen_senden(zeile[:bland_arbeit]);  feld_vor(1)
    if zeile[:berufsgruppe] == "Azubi"
      zeichen_senden('{DOWN}')
    elsif zeile[:berufsgruppe] == "sozialvers.freier GGF"
      zeichen_senden('{DOWN}')
      zeichen_senden('{DOWN}')
      @feld_minijob_aktiv = false
      @feld_kinderlos_aktiv = false
    end
    feld_vor(1) # defaultbelegung "angestellte/arbeiter"
    if zeile[:durchfuehrungsweg] == "Pensionskasse"
      zeichen_senden('{DOWN}')
    elsif zeile[:durchfuehrungsweg] == "Unterstützungskasse"
      zeichen_senden('{DOWN}')
      zeichen_senden('{DOWN}')
      @feld_pauschalverst_aktiv = false
    end
    feld_vor(1) #defaultbelegung "Direktversicherung"
    if @feld_pauschalverst_aktiv == true
      zeichen_senden(' ') if zeile[:pausch_steuer40b] == "j"
      feld_vor(1)
    end
    if @feld_minijob_aktiv == true
      zeichen_senden(' ') if zeile[:minijob_ok] == "ja"
      feld_vor(1)
    end
    if @feld_kinderlos_aktiv == true
      zeichen_senden(' ') if zeile[:kinderlos] == "j"
      feld_vor(1)
    end
    feld_vor(5)
 
    #Blatt 2
    zeichen_senden('{RIGHT}')
    feld_zurueck(6)
    #anfang des 2 Blattes
    zeichen_senden('{TAB}') #umwandlung von bestandteilen des einkommens
    feld_vor(1) #erst netto / brutto feld auswaehlen
    if zeile[:verzicht_als_netto] == "brutto"
      feld_vor(1);  zeichen_senden(' ')
    else
      zeichen_senden(' ');  feld_zurueck(1)
    end
    zeichen_senden(zeile[:verzicht_betrag])
    feld_vor(5)

    #Blatt 3
    zeichen_senden('{RIGHT}')
    feld_zurueck(5)
    #anfang des 3 Blattes
    zeichen_senden(zeile[:vl_arbeitgeber]);  feld_vor(1)
    zeichen_senden('{TAB}'); feld_vor(1) #ueberweisung vl
    zeichen_senden(' ') if zeile[:vl_als_beitrag] == "nein"
    feld_vor(3)

    #Blatt 4
    zeichen_senden('{RIGHT}')
    if zeile[:ag_zuschuss] == 0 || zeile[:ag_zuschuss] == nil
      feld_zurueck(2)
      #ergebnis-button des 4 blattes
    else
      feld_zurueck(3)
      #anfang des 4 blattes
      zeichen_senden(' ')
      if zeile[:ag_zuschuss_als_absolut] == "%"
        feld_zurueck(1)
        zeichen_senden(' ')
      end
      zeichen_senden(zeile[:ag_zuschuss]); feld_vor(1)
    end
    berechnung_starten
  end

  def berechnung_starten #besser waere es, wenn der button "ergebnis" direkt angesprochen werden kann
    eingabe_bestaetigen
    sleep(0.1)
    eingabe_bestaetigen
  end

  def vb_senden(vb_abfrage)
    @xlapp.Run "#{@dateiname}!#{vb_abfrage}"
  end

  def maske_schliessen #ueber button "schliessen" siehe kommentar "berechnung_starten"
    @xlapp.ActiveWorkbook.Close
  end

  def excel_beenden
    @excel_controller.quit_excel
  end

end

