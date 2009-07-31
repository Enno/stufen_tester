
require File.dirname(File.dirname(__FILE__)) +  '/lib/excel_leser.rb'


describe ExcelLeser do
  before(:each) do
    mappen_name = "Tabelle-4sr_a.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    @el = ExcelLeser.new(mappen_pfad + mappen_name, "Global", "Tabelle")
  end
  after(:each) do
    @el.excel_beenden
  end

  it "sollte existieren" do
    @el.should_not be_nil
  end

  it "sollte auf die Methode 'zeile' reagieren und einen Hash zurückgeben" do
    @el.zeile(22).is_a?(Hash).should be_true
  end

  it "sollte Zeile 22 korrekt einlesen" do
    z22 = @el.zeile(22)

    z22[:name].should                  == "Hans Meier"
    z22[:verzicht_betrag].should       == 50.0
  end

  it "sollte Zeile 21 korrekt einlesen" do
    z21 = @el.zeile(21)

    z21[:name][0..6].should                  == "Gerda M" #üller"
    #@el.wert(22, :name).should                  == "Hans Meier"
    z21[:berufsgruppe].should          == "Angestellte/Arbeiter"
  end

  it "sollte ausreichend Spalten enthalten" do
    @el.zeile(21).size.should >= SPALTEN_UEBERSCHRIFTEN.size
    @el.zeile(22).size.should >= SPALTEN_UEBERSCHRIFTEN.size
    @el.zeile(23).size.should >= SPALTEN_UEBERSCHRIFTEN.size
  end
end

