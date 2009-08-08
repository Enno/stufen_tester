

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
    @masken_controller = TastenSender.new(:wartezeit => 0.2)
 
    @feld_kinderlos_aktiv = true
    @feld_pauschalverst_aktiv = true
    @feld_minijob_aktiv = true
    case @xlapp.version
    when "12.0"
      @fenstername = 'Microsoft Excel' #fuer office 07 anwendungen
    when "11.0"
      @fenstername = "Microsoft Excel - #{@dateiname}" # für Office XP/2002
    end
  end

  def maske_oeffnen
    #@xlapp.Run "#{@datei_name}!#{@proc_name}"
    @masken_controller.sende_tasten(@fenstername, "%{F8}#{@proc_name}%{a}", :wartezeit => 0.2, :fenster_fehlt=>"Komischerweise fehlt das Excel-Fenster")
    sleep(0.5)
  end

  def feld_vor(anzahl)
    zu_sendende_tabs = "{TAB}" * anzahl
    tasten_senden("#{zu_sendende_tabs}", :wartezeit => 0.01 )
  end

  def feld_zurueck(anzahl)
    zu_sendende_tabs = "+{TAB}" * anzahl
    tasten_senden("#{zu_sendende_tabs}", :wartezeit => 0.01)
  end

  def eingabe_bestaetigen
    tasten_senden( "{ENTER}") #, :wartezeit => 0.1)
  end

  def tasten_senden(zeichen, optionen = {})
    @masken_controller.sende_tasten(@fenstername, "#{zeichen}", optionen)
  end

  @@register_karten = [
    {:felder => [ :name,
        :bruttogehalt,
        :freibetrag,
        {:k_vers_art => ["g","p"]},
        :steuerklasse,
        :kinder_fb,
        :kirchensteuer,
        :bland_wohnsitz,
        :bland_arbeit,
        :berufsgruppe,
        :durchfuehrungsweg,
        :pausch_steuer40b,
        {:minijob_ok => false},
        :kinderlos
      ],
      :anzahl_zurueck => 5
    }

  ]

  def wert_eintragen_fuer(datensatz, symbol_oder_hash)
    case symbol_oder_hash
    when Symbol
      art = :direkt
      sym = symbol_oder_hash
    when Hash
      rechte_seite = symbol_oder_hash.values.first
      art = rechte_seite.is_a?(Array) ? :radio_group : :checkbox
      sym = symbol_oder_hash.keys.first
    end

    einzutragender_wert = datensatz[sym]
    case art
    when :direkt
      einzutragender_wert.is_a?(Float) ? tasten_senden(dezimalzahl_fuer_office_umwandeln(einzutragender_wert)) : tasten_senden(einzutragender_wert)
      feld_vor(1)
    when :checkbox
      vorbelegung = rechte_seite
      tasten_senden(' ') if vorbelegung ^ einzutragender_wert # exclusive or
      feld_vor(1)
    when :radio_group
      auswahl_liste = rechte_seite
      auswahl_liste.each do |moegl_wert|
        tasten_senden(' ') if moegl_wert == einzutragender_wert
        feld_vor(1)
      end
    end
  end

  def dezimalzahl_fuer_office_umwandeln(einzutragender_wert)
    return einzutragender_wert.to_s.gsub(/[.]/, ',')
  end

  def maske_fuellen(datensatz)
    maske_oeffnen

    letzte_anzahl_zurueck = nil
    @@register_karten.each do |karten_beschreibung|
      felder_in_karte = karten_beschreibung[:felder]
      felder_in_karte.each do |feld_info|
        wert_eintragen_fuer(datensatz, feld_info)
      end
      letzte_anzahl_zurueck = karten_beschreibung[:anzahl_zurueck]
    end
    feld_vor(15)
    feld_zurueck(letzte_anzahl_zurueck)
    berechnung_starten
  end



  def xxx_maske_fuellen(zeile)
    maske_oeffnen
    #Blatt 1
    tasten_senden(zeile[:name]);  feld_vor(1)
    tasten_senden(zeile[:bruttogehalt]);  feld_vor(1)
    tasten_senden(zeile[:freibetrag]);  feld_vor(1)
    if zeile[:k_vers_art] == "p"
      feld_vor(1);  tasten_senden(' ')
      @feld_kinderlos_aktiv = false
    else
      tasten_senden(' ');  feld_vor(1)
    end 
    feld_vor(1)
    tasten_senden(zeile[:steuerklasse]);  feld_vor(1)
    tasten_senden(zeile[:kinder_fb]);  feld_vor(1)
    tasten_senden(' ') if zeile[:kirchensteuer] == "n"
    feld_vor(1)
    tasten_senden(zeile[:bland_wohnsitz]);  feld_vor(1)
    tasten_senden(zeile[:bland_arbeit]);  feld_vor(1)
    if zeile[:berufsgruppe] == "Azubi"
      tasten_senden('{DOWN}')
    elsif zeile[:berufsgruppe] == "sozialvers.freier GGF"
      tasten_senden('{DOWN}')
      tasten_senden('{DOWN}')
      @feld_minijob_aktiv = false
      @feld_kinderlos_aktiv = false
    end
    feld_vor(1) # defaultbelegung "angestellte/arbeiter"
    if zeile[:durchfuehrungsweg] == "Pensionskasse"
      tasten_senden('{DOWN}')
    elsif zeile[:durchfuehrungsweg] == "Unterstützungskasse"
      tasten_senden('{DOWN}')
      tasten_senden('{DOWN}')
      @feld_pauschalverst_aktiv = false
    end
    feld_vor(1) #defaultbelegung "Direktversicherung"
    if @feld_pauschalverst_aktiv == true
      tasten_senden(' ') if zeile[:pausch_steuer40b] == "j"
      feld_vor(1)
    end
    if @feld_minijob_aktiv == true
      tasten_senden(' ') if zeile[:minijob_ok] == "ja"
      feld_vor(1)
    end
    if @feld_kinderlos_aktiv == true
      tasten_senden(' ') if zeile[:kinderlos] == "j"
      feld_vor(1)
    end
    feld_vor(5)
 
    #Blatt 2
    tasten_senden('{RIGHT}')
    feld_zurueck(6)
    #anfang des 2 Blattes
    tasten_senden('{TAB}') #umwandlung von bestandteilen des einkommens
    feld_vor(1) #erst netto / brutto feld auswaehlen
    if zeile[:verzicht_als_netto] == "brutto"
      feld_vor(1);  tasten_senden(' ')
    else
      tasten_senden(' ');  feld_zurueck(1)
    end
    tasten_senden(zeile[:verzicht_betrag])
    feld_vor(5)

    #Blatt 3
    tasten_senden('{RIGHT}')
    feld_zurueck(5)
    #anfang des 3 Blattes
    tasten_senden(zeile[:vl_arbeitgeber]);  feld_vor(1)
    tasten_senden('{TAB}'); feld_vor(1) #ueberweisung vl
    tasten_senden(' ') if zeile[:vl_als_beitrag] == "nein"
    feld_vor(3)

    #Blatt 4
    tasten_senden('{RIGHT}')
    if zeile[:ag_zuschuss] == 0 || zeile[:ag_zuschuss] == nil
      feld_zurueck(2)
      #ergebnis-button des 4 blattes
    else
      feld_zurueck(3)
      #anfang des 4 blattes
      tasten_senden(' ')
      if zeile[:ag_zuschuss_als_absolut] == "%"
        feld_zurueck(1)
        tasten_senden(' ')
      end
      tasten_senden(zeile[:ag_zuschuss]); feld_vor(1)
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

