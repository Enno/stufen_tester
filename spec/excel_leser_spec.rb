
require 'win32ole'
require File.dirname(File.dirname(__FILE__)) +  '/lib/excel_leser.rb'
require 'spec'

describe ExcelLeser do
  before(:each) do
    mappen_name = "Tabelle-4sr_a.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    @el = ExcelLeser.new(mappen_pfad + mappen_name, "Global", "Tabelle")
  end
  after(:each) do
    @el.quit_excel
  end

  it "sollte existieren" do
    @el.should_not be_nil
  end

  it "sollte auf die Methode 'zeile' reagieren" do
    @el.zeile(22).should_not be_nil
  end

  it "sollte bei 'zeile' einen Hash zurückgeben" do
    @el.zeile(22).is_a?(Hash).should be_true
  end

  it "sollte Zeile 22 korrekt einlesen" do
    z22 = @el.zeile(22)

    z22[:name].should                  == "Hans Meier"
    z22[:verzicht_als_netto].should    == 20.0
  end

  it "sollte Zeile 21 korrekt einlesen" do
    z21 = @el.zeile(21)

    z21[:name].should                  == "Gerda Müller"
    z21[:berufsgruppe].should          == "Angestellte/Arbeiter"
  end
end

