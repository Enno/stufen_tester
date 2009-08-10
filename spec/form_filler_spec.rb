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

  it "sollte Namen korrekt eintragen" do
    zeile = {:name => "Max Peter", :bruttogehalt => 2000, :k_vers_art => "p",      :steuerklasse => "III"}
    @ff.maske_fuellen(zeile)
    @ff.vb_senden("Abfrage_Feld_name").should == "Max Peter"
  end

  it "sollte Kinderfreibetrag korrekt eintragen" do
    @ff.maske_fuellen(:bruttogehalt => 2000,
      :k_vers_art => "g",
      :kinder_fb => 2,
      #:kinderlos => "j",
      :verzicht_betrag => 30,
      :verzicht_als_netto => "brutto")
    @ff.vb_senden("Abfrage_Feld_kinderfreibetraege").should == 2
  end

  it "sollte Verzichts-Betrag korrekt eintragen" do
    betrag = 43
    @ff.maske_fuellen(:verzicht_als_netto => "brutto",
      :verzicht_betrag => betrag,
      :bruttogehalt => 2000,
      :ag_zuschuss => 20,
      :ag_zuschuss_als_absolut => "€")
    @ff.vb_senden("Abfrage_Feld_nvz_betrag").should == 43
#    @ff.vb_senden("Abfrage_Feld_AG_Zuschuss").should == true
#    @ff.vb_senden("Abfrage_Feld_ag_prozent").should == false
#    @ff.vb_senden("Abfrage_Feld_AG_Beitrag").should == 20
  end

  it "sollte Kommazahlen korekt eintragen" do
    brutto_betrag = 2000.50
    kfb = 2.5
    zeile = {:name => "Max Peter", :bruttogehalt => brutto_betrag,
             :kinder_fb => kfb
    }
    @ff.maske_fuellen(zeile)
    @ff.vb_senden("Abfrage_Feld_gehalt").should == brutto_betrag
    @ff.vb_senden("Abfrage_Feld_kinderfreibetraege").should == kfb
  end

  it "sollte MinijobOK korrekt eintragen" do
    @ff.maske_fuellen(:minijob_ok => true, :bruttogehalt => 2000)
    @ff.vb_senden("Abfrage_Feld_Minijob").should == true
  end

=begin
=end
end

