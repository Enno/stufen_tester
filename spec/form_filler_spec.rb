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
    2.times { @ff.send_keys("%{F11}") } # Workaround für nicht funktionierende Tab-Tasten
  end

  after(:each) do
    @ff.close_template
    @ff.quit_excel
  end
  #
  #  it "sollte Namen korrekt eintragen" do
  #    zeile = {:name => "Max Peter", :bruttogehalt => 2000, :k_vers_art => "p",
  #      :steuerklasse => "III", :bland_arbeit => "Berlin-West"}
  #    @ff.populate_template(zeile)
  #    @ff.vb_send("Abfrage_Feld_name").should == "Max Peter"
  #  end
  #
  #  it "sollte Kinderfreibetrag korrekt eintragen" do
  #    @ff.populate_template(:bruttogehalt => 2000,
  #      :k_vers_art => "g",
  #      :kinder_fb => 2,
  #      #:kinderlos => "j",
  #      :verzicht_betrag => 30,
  #      :verzicht_als_netto => "brutto")
  #    @ff.vb_send("Abfrage_Feld_kinderfreibetraege").should == 2
  #  end
  #
  #  it "sollte Verzichts-Betrag korrekt eintragen" do
  #    betrag = 43
  #    @ff.populate_template(:verzicht_als_netto => "brutto",
  #      :verzicht_betrag => betrag,
  #      :bruttogehalt => 2000,
  #      :ag_zuschuss => 20,
  #      :ag_zuschuss_als_absolut => "€")
  #    @ff.vb_send("Abfrage_Feld_nvz_betrag").should == 43
  ##    @ff.vb_send("Abfrage_Feld_AG_Zuschuss").should == true
  ##    @ff.vb_send("Abfrage_Feld_ag_prozent").should == false
  ##    @ff.vb_send("Abfrage_Feld_AG_Beitrag").should == 20
  #  end
  #
  #  it "sollte Kommazahlen korekt eintragen" do
  #    brutto_betrag = 2000.50
  #    kfb = 2.5
  #    zeile = {:name => "Max Peter", :bruttogehalt => brutto_betrag,
  #             :kinder_fb => kfb
  #    }
  #    @ff.populate_template(zeile)
  #    @ff.vb_send("Abfrage_Feld_gehalt").should == brutto_betrag
  #    @ff.vb_send("Abfrage_Feld_kinderfreibetraege").should == kfb
  #  end
  #
  #  it "sollte MinijobOK korrekt eintragen" do
  #    @ff.populate_template(:minijob_ok => true, :bruttogehalt => 2000)
  #    @ff.vb_send("Abfrage_Feld_Minijob").should == true
  #  end
  #
  it "sollte für vollen Datensatz funktionieren" do
    datensatz = {
      :name                   => "Gerda Schulze",
      ##:personal_nr=> 1.0,
      ##:geb_datum=>"1966/05/02 00:00:00",
      :bruttogehalt           => 2000.0,
      :freibetrag             => 20,
      :k_vers_art             => "g",
      :steuerklasse           => "II",
      :kinder_fb              => 0.5,
      :kirchensteuer          => true, #,
      #:bland_wohnsitz=>"Berlin-Ost",
      #:bland_arbeit=>"Berlin-West",
      :berufsgruppe           => "Angestellte/Arbeiter",
      :durchfuehrungsweg      => "Direktversicherung",
      :pausch_steuer40b       => true,
      :minijob_ok             => false,
      :kinderlos              => true,

#      :nvz                    => true,
      :verzicht_betrag        => 57.57,
      :verzicht_als_netto     => false,

      :vl_arbeitgeber         => 40.0,
      :vl_arbeitnehmer        => 0.0,

      :vl_als_beitrag         => true,

 #     :ag_zuschuss_ok         => true,
      :ag_zuschuss_als_absolut=> true, #"€",
      :ag_zuschuss            => 20
    }
    @ff.populate_template datensatz
    keys_zu_stufenrechner_namen = {
      :name                    => "name",
      :bruttogehalt            => "gehalt",
      :freibetrag              => "Freibetrag",
      #"kv_pflicht",
      #"KV_privat",
      #    :steuerklasse       => "Steuerklasse", # problem: roemische und lateinische ziffern
      :kinder_fb               => "kinderfreibetraege",
      :kirchensteuer           => "Kirchensteuer",
      #    :bland_wohnsitz     => "Wohnsitz",
      #    :bland_arbeit       => "arbeitsstaette",
      :berufsgruppe            => "Berufsgruppe",
      :durchfuehrungsweg       => "bavweg",
      :pausch_steuer40b        => "dive_40b_vorhanden",
      :minijob_ok              => "Minijob",
      :kinderlos               => "erh_pvsatz",

      #      :nvz            => "nvz",
      :verzicht_betrag         => "nvz_betrag",
      :verzicht_als_netto      => "nvz_netto",
      #      :verzicht_als_netto      => "nvz_brutto",
      :vl_arbeitgeber          => "VL_AG",
      :vl_arbeitnehmer         => "VL_AN",
      #"VL_gesamt",
      :vl_als_beitrag          => "vl",
      ##"kv_satz_durchschn",
      ##"kv_satz_indiv_satz",
      ##"KV_Satz",
      ##"kv_wechsel",
      ##"kv_satz_neu",
#      :ag_zuschuss_ok          => "AG_Zuschuss",
      :ag_zuschuss             => "AG_Beitrag",
      :ag_zuschuss_als_absolut => "ag_betrag", #"ag_prozent",
      ##"vetrieb",
      #"pv_pflicht"
    }
    keys_zu_stufenrechner_namen.each do |key, sr_name|
      [key, @ff.vb_send("Abfrage_Feld_#{sr_name}")].should == [key, datensatz[key]]
    end
    puts "#{keys_zu_stufenrechner_namen.size} felder getestet"

  end
=begin
=end
end

