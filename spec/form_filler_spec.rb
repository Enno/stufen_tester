# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'form_filler'

#require 'tasten_sender'

describe FormFiller do
  before(:each) do
    mappen_name = "sr38a_entkernt_test.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    # dateiname = mappen_pfad + mappen_name
    start_proc_name = "Entgeltumwandlungsrechner_starten"
    @ff = FormFiller.new(mappen_pfad, mappen_name, start_proc_name)
  end

  after(:each) do
    @ff.maske_schliessen
    @ff.excel_beenden
  end



  it "sollte Namen korrekt eintragen" do
    zeile = {:name => "Max Peter", :bruttogehalt => 2000, :k_vers_art => "p",
      :steuerklasse => "II"}
    @ff.maske_fuellen(zeile)
    #TastenSender.new().sende_tasten('Microsoft Excel', nil).should == true

    #    @ff.ergebnis_anfordern
    @ff.vb_senden("Abfrage_Feld_name").should == "Max Peter"
  end

  it "sollte Kinderfreibetrag korrekt eintragen" do
    @ff.maske_fuellen(:bruttogehalt => 2000, :k_vers_art => "g", 
      :kinder_fb => 2, :kinderlos => "j")
    @ff.vb_senden("Abfrage_Feld_kinderfreibetraege").should == 2
  end
  #
  #  it "sollte Verzichts-Betrag korrekt eintragen" do
  #    betrag = 43.50
  #    @ff.maske_fuellen(:verzicht_betrag => betrag, :bruttogehalt => 2000)
  #    @ff.vb_senden("Abfrage_Feld_nvz_betrag").should == 43.50
  #  end
  #
  #  it "sollte MinijobOK korrekt eintragen" do
  #    @ff.maske_fuellen(:minijob_ok => true, :bruttogehalt => 2000)
  #    @ff.vb_senden("Abfrage_Feld_Minijob").should == true
  #  end
end

