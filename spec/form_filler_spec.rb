# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'form_filler'

require 'tasten_sender'

describe FormFiller do
  before(:each) do
    mappen_name = "sr38a_entkernt_test.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    start_proc_name = "Entgeltumwandlungsrechner_starten"
    @ff = FormFiller.new(mappen_pfad, mappen_name, start_proc_name)
    2.times { @ff.tasten_senden("%{F11}") } # Workaround für nicht funktionierende Tab-Tasten
  end

  after(:each) do
    @ff.maske_schliessen
    @ff.excel_beenden
  end
#
#  it "sollte Namen korrekt eintragen" do
#    zeile = {:name => "Max Peter", :bruttogehalt => 2000, :k_vers_art => "p",
#      :steuerklasse => "III", :bland_arbeit => "Berlin-West"}
#    @ff.maske_fuellen(zeile)
#    @ff.vb_senden("Abfrage_Feld_name").should == "Max Peter"
#  end
#
#  it "sollte Kinderfreibetrag korrekt eintragen" do
#    @ff.maske_fuellen(:bruttogehalt => 2000,
#      :k_vers_art => "g",
#      :kinder_fb => 2,
#      #:kinderlos => "j",
#      :verzicht_betrag => 30,
#      :verzicht_als_netto => "brutto")
#    @ff.vb_senden("Abfrage_Feld_kinderfreibetraege").should == 2
#  end
#
#  it "sollte Verzichts-Betrag korrekt eintragen" do
#    betrag = 43
#    @ff.maske_fuellen(:verzicht_als_netto => "brutto",
#      :verzicht_betrag => betrag,
#      :bruttogehalt => 2000,
#      :ag_zuschuss => 20,
#      :ag_zuschuss_als_absolut => "€")
#    @ff.vb_senden("Abfrage_Feld_nvz_betrag").should == 43
##    @ff.vb_senden("Abfrage_Feld_AG_Zuschuss").should == true
##    @ff.vb_senden("Abfrage_Feld_ag_prozent").should == false
##    @ff.vb_senden("Abfrage_Feld_AG_Beitrag").should == 20
#  end
#
#  it "sollte Kommazahlen korekt eintragen" do
#    brutto_betrag = 2000.50
#    kfb = 2.5
#    zeile = {:name => "Max Peter", :bruttogehalt => brutto_betrag,
#             :kinder_fb => kfb
#    }
#    @ff.maske_fuellen(zeile)
#    @ff.vb_senden("Abfrage_Feld_gehalt").should == brutto_betrag
#    @ff.vb_senden("Abfrage_Feld_kinderfreibetraege").should == kfb
#  end
#
#  it "sollte MinijobOK korrekt eintragen" do
#    @ff.maske_fuellen(:minijob_ok => true, :bruttogehalt => 2000)
#    @ff.vb_senden("Abfrage_Feld_Minijob").should == true
#  end
#
  it "sollte für vollen Datensatz funktionieren" do
    datensatz = {:name=>"Gerda Schulze", #:geb_datum=>"1966/05/02 00:00:00",
      :vl_als_beitrag=>true, :bruttogehalt=>2000.0, #:bland_arbeit=>"Berlin-West",
      #:bland_wohnsitz=>"Berlin-Ost",
      :kinderlos=>false, :freibetrag=>nil,
      :kirchensteuer=>true, #, :verzicht_als_netto=>"netto"
      :berufsgruppe=>"Angestellte/Arbeiter", :verzicht_betrag=>57.57,
      :ag_zuschuss_als_absolut=>"%", :k_vers_art=>"g", :pausch_steuer40b=>true,
      :minijob_ok=>false, :personal_nr=>1.0, :steuerklasse=>"I",
      :kinder_fb=>0.5,
      :durchfuehrungsweg=>"Direktversicherung",
      :vl_arbeitgeber=>40.0,
      :vl_arbeitnehmer=>0.0,
      :ueberweisungvl_keine_ahnung_welches_feld =>40,
      :ag_zuschuss_ok => true,
      :ag_zuschuss => 20
    }
    @ff.maske_fuellen datensatz
    keys_zu_stufenrechner_namen = {
      :name => "name",
      :bruttogehalt=>"gehalt",
#    :freibetrag=>"Freibetrag",
#    :bland_wohnsitz=>"Wohnsitz",
#    :bland_arbeit=>"arbeitsstaette",
#    :steuerklasse=>"Steuerklasse",
    :kinder_fb=>"kinderfreibetraege",
    :kirchensteuer=>"Kirchensteuer",
    :berufsgruppe=>"Berufsgruppe",
    #"nvz",
    :verzicht_betrag=>"nvz_betrag",
    #"nvz_netto",
    #"nvz_brutto",
    #"vl",
    :vl_arbeitgeber=>"VL_AG",
    :vl_arbeitnehmer=>"VL_AN",
    #"VL_gesamt",
    ##"kv_satz_durchschn",
    ##"kv_satz_indiv_satz",
    #"kv_pflicht",
    #"KV_privat",
    ##"KV_Satz",
    ##"kv_wechsel",
    ##"kv_satz_neu",
#    :ag_zuschuss=>"AG_Zuschuss",
    #"AG_betrag",
    #"ag_prozent",
    #"AG_Beitrag",
    #"bavweg",
    ##"vetrieb",
    :minijob_ok=>"Minijob",
#    :pausch_steuer40b=>"dive_40b_vorhanden",
    ##"erh_pvsatz",
    #"pv_pflicht"
       }
    keys_zu_stufenrechner_namen.each do |key, sr_name|
      [key, @ff.vb_senden("Abfrage_Feld_#{sr_name}")].should == [key, datensatz[key]]
    end
    puts "#{keys_zu_stufenrechner_namen.size} felder getestet"

  end
=begin
=end
end

