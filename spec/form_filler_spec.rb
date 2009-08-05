# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'form_filler'

describe FormFiller do
  before(:each) do
    mappen_name = "sr38a_entkernt.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    dateiname = mappen_pfad + mappen_name
    start_proc_name = "Entgeltumwandlungsrechner_starten"
    @ff = FormFiller.new(dateiname, start_proc_name)
  end

  after(:each) do
    @el.ergebnis_anfordern
  end


  it "sollte nicht abstürzen" do
    zeile = {:name=>"Max Peter"}
    @ff.fill(zeile)
  end
end

