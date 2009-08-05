# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'form_filler'

#require 'tasten_sender'

describe FormFiller do
  before(:each) do
    mappen_name = "sr38a_entkernt.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    dateiname = mappen_pfad + mappen_name
    start_proc_name = "Entgeltumwandlungsrechner_starten"
    @ff = FormFiller.new(dateiname, start_proc_name)
  end

  after(:each) do
    @ff.excel_beenden
  end

  it "sollte nicht abstürzen" do
    name = "Max Peter"
    zeile = {:name => name}
    @ff.fill(zeile)
  end

  it "sollte Namen korrekt eintragen" do
    name = "Max Peter"
    zeile = {:name => name, :bruttogehalt => 2000}
    @ff.fill(zeile)
    #TastenSender.new().sende_tasten('Microsoft Excel', nil).should == true

#    @ff.ergebnis_anfordern
    erg = @ff.vb_senden("Abfrage_Feld_Name")
    erg.should == name
  end

  it "sollte Kinderfreibetrag korrekt eintragen" do
    @ff.fill(:kinder_fb => 2, :bruttogehalt => 2000)
    @ff.vb_senden("Abfrage_Feld_kinderfreibetraege").should == 2
  end

  it "sollte Verzichts-Betrag korrekt eintragen" do
    betrag = 43.50
    @ff.fill(:verzicht_betrag => betrag, :bruttogehalt => 2000)
    @ff.vb_senden("Abfrage_Feld_nvz_betrag").should == 43.50
  end

  it "sollte MinijobOK korrekt eintragen" do
    @ff.fill(:minijob_ok => true, :bruttogehalt => 2000)
    @ff.vb_senden("Abfrage_Feld_Minijob").should == true
  end
end

