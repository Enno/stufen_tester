

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
    @automatisch_auf_naechstem_feld = false
    @rueckwaerts_in_die_zellen_eintragen = false
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

  def feld_vor(anzahl = 1)
    shift_code = anzahl < 0 ? "+" : ""
    zu_sendende_tabs = "#{shift_code}{TAB}" * anzahl.abs
    tasten_senden("#{zu_sendende_tabs}", :wartezeit => 0.702)
  end

  def feld_zurueck(anzahl)
    feld_vor(-anzahl)
  end

  def eingabe_bestaetigen
    tasten_senden("{ENTER}", :wartezeit => 0.2)
  end

  def tasten_senden(zeichen, optionen = {})
    @masken_controller.sende_tasten(@fenstername, "#{zeichen}", optionen)
  end

  @@register_karten = [
    {:felder_blatt1 => [
        :name,
        :bruttogehalt
#        :freibetrag,
#        {:k_vers_art => ["g","p"]},
#        :steuerklasse,
#        :kinder_fb,
#        {:kirchensteuer => true},
#        :bland_wohnsitz,
#        :bland_arbeit,
#        :berufsgruppe,
#        :durchfuehrungsweg,
#        {:pausch_steuer40b => false},
#        {:minijob_ok => false},
#        {:kinderlos => false},
      ],
      #:automatisch_auf_naechstem_feld => false,
      #:rueckwaerts_eintragen => false
    },
    {:felder_blatt2 => [
        {:umwandlgvon_keine_ahnung_welches_feld => false},
        :verzicht_betrag,
        {:verzicht_als_netto => ["netto", "brutto"]}
      ]
      #      :automatisch_auf_naechstem_feld => false,
      #      :rueckwaerts_eintragen => false
    },
    {:felder_blatt3 => [
        :vl_arbeitgeber,
        :ueberweisungvl_keine_ahnung_welches_feld,
        {:vl_als_beitrag => true}
      ]
      #      :automatisch_auf_naechstem_feld => false,
      #      :rueckwaerts_eintragen => false
    },
    {:felder_blatt4 => [
        {:ag_zuschuss_ok => {:vorbelegung  => false, 
            :sprung_korrektur => -2,
            :macht_aktiv => [:ag_zuschuss, :ag_zuschuss_als_absolut]}
        },
        {:ag_zuschuss_als_absolut => ["€", "%"]},
        :ag_zuschuss
      ],
#      :sprung_korrektur => {:ag_zuschuss_ok => -2},
#      :macht_aktiv => {:ag_zuschuss_ok => [:ag_zuschuss, {:bedingung => proc{|wert| wert == "k"}}]}
      #:automatisch_auf_naechstem_feld => true,
      #:rueckwaerts_eintragen => true
    }
  ]

  def wert_eintragen_fuer(datensatz, symbol_oder_hash)
    case symbol_oder_hash
    when Symbol
      art = :direkt
      sym = symbol_oder_hash
    when Hash
      rechte_seite = symbol_oder_hash.values.first
      art = case rechte_seite
      when Array     then :radio_group
      when true      then :checkbox
      when false     then :checkbox
      when Hash      then :komplex
      end
      sym = symbol_oder_hash.keys.first
    end

    return  if @inaktive_felder.include? sym

    einzutragender_wert = datensatz[sym]
    case art
    when :direkt
      einzutragender_wert.is_a?(Float) ? tasten_senden(dezimalzahl_fuer_office_umwandeln(einzutragender_wert)) : tasten_senden(einzutragender_wert)
      feld_vor
      #@rueckwaerts_in_die_zellen_eintragen ? feld_zurueck(1) : feld_vor(1)
    when :checkbox
      #      if rechte_seite.is_a?(Symbol)
      #        vorbelegung = false
      #        if ((datensatz[rechte_seite] == 0) || (datensatz[rechte_seite] == nil)) then
      #          einzutragender_wert = false
      #          puts "bin hier"
      #          puts vorbelegung, einzutragender_wert
      #          @automatisch_auf_naechstem_feld = false
      #        else
      #          einzutragender_wert = true
      #          @automatisch_auf_naechstem_feld = true
      #        end
      #      else
      #        vorbelegung = rechte_seite
      #      end
      vorbelegung = rechte_seite
      tasten_senden(' ') if vorbelegung ^ einzutragender_wert # exclusive or
      feld_vor 
      #      if @automatisch_auf_naechstem_feld == false
      #        @rueckwaerts_in_die_zellen_eintragen ? feld_zurueck(1) : feld_vor(1)
      #      end
    when :komplex
      vorbelegung = rechte_seite[:vorbelegung]
      tasten_senden(' ') if vorbelegung ^ einzutragender_wert # exclusive or
      feld_vor( rechte_seite[:sprung_korrektur] ?  rechte_seite[:sprung_korrektur] : 1)
      neue_inaktive = (einzutragender_wert ? [] : rechte_seite[:macht_aktiv])
      #neue_inaktive = (rechte_seite[:macht_aktiv] if einzutragender_wert)
      @inaktive_felder += neue_inaktive if neue_inaktive
    when :radio_group
      auswahl_liste = rechte_seite
      auswahl_liste.each do |moegl_wert|
        tasten_senden(' ') if moegl_wert == einzutragender_wert
        feld_vor
        #@rueckwaerts_in_die_zellen_eintragen ? feld_zurueck(1) : feld_vor(1)
      end
    end
  end

  def dezimalzahl_fuer_office_umwandeln(einzutragender_wert)
    return einzutragender_wert.to_s.gsub(/[.]/, ',')
  end

  def maske_fuellen(datensatz)
    maske_oeffnen
    register_karten_index = 1
    @@register_karten.each do |karten_beschreibung|
      felder_des_aktuellen_blattes = "felder_blatt#{register_karten_index}".to_sym
      felder_in_karte = karten_beschreibung[felder_des_aktuellen_blattes]
      #@automatisch_auf_naechstem_feld = karten_beschreibung[:automatisch_auf_naechstem_feld]
      #@rueckwaerts_in_die_zellen_eintragen = karten_beschreibung[:rueckwaerts_eintragen]
      @inaktive_felder = []
      felder_in_karte.each do |feld_info|
        wert_eintragen_fuer(datensatz, feld_info)
      end
      feld_vor(25)
      break if register_karten_index == 4
      tasten_senden('{RIGHT}')
      feld_zurueck(25)
      register_karten_index += 1
    end
    #feld_zurueck(1)
    feld_zurueck(1)
    feld_zurueck(1)
    berechnung_starten
  end

  def berechnung_starten #besser waere es, wenn der button "ergebnis" direkt angesprochen werden kann
    sleep 2
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

