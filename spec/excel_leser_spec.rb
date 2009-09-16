
require File.dirname(File.dirname(__FILE__)) +  '/lib/excel_leser.rb'


describe ExcelLeser, "ohne reale Daten" do
  before(:each) do
    dummy_pfad = "mappen_pfad + mappen_name"
    @ec = mock("MockExcelController")
    ExcelController.stub!(:new => @ec)
    @ec.should_receive(:open_excel_file)
    @ec.should_receive(:excel_appl)
    @el = ExcelLeser.new(dummy_pfad, "Global", "Tabelle")
  end

  it "should convert integer Floats to Integers" do
    eingabe_array = [1.0, "abc", -42.0, 100.0]
    ausgabe_array = eingabe_array.clone
    @el.werte_auf_integer_pruefen(ausgabe_array)
    ausgabe_array.should == eingabe_array
    ausgabe_array[0].should be_a(Integer)
    ausgabe_array[2].should be_a(Integer)
    ausgabe_array[3].should be_a(Integer)
  end

  it "should convert only integer Floats to Integers" do
    eingabe_array = [1.00001, -42, "hallo", 100.0000300]
    ausgabe_array = eingabe_array.clone
    @el.werte_auf_integer_pruefen(ausgabe_array)
    ausgabe_array.should == eingabe_array
    ausgabe_array[0].should be_a(Float)
    ausgabe_array[1].should be_a(Integer)
    ausgabe_array[2].should be_a(String)
    ausgabe_array[3].should be_a(Float)
  end
end


describe ExcelLeser, "mit realen Daten" do
  before(:each) do
    mappen_name = "test.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    @el = ExcelLeser.new(mappen_pfad + mappen_name, "Global", "Tabelle")
  end
  after(:each) do
    @el.excel_beenden
  end

  #=begin
  it "sollte existieren" do
    @el.should_not be_nil
  end

  it "sollte interne Methode zeilen_array korrekt ausführen" do
    ues = @el.zeile_als_array(19)
    ues.should be_a Array
    ues.first.should == "Name, Vorname"
  end

  it "sollte auf die Methode 'zeile' reagieren und einen Hash zurückgeben" do
    @el.zeile(22).is_a?(Hash).should be_true
  end

  it "sollte Zeile 22 korrekt einlesen" do
    z22 = @el.zeile(22)

    z22[:name].should                                              == "Hans Meier"
    z22[:verzicht_betrag].should                                   == 50.0
    z22[:akt_gehaltsabr_monatl_brutto_gehalt].should               == 2000.00
    z22[:akt_gehaltsabr_ag_anteil_vl].should                       == 40.00
    z22[:akt_gehaltsabr_beitrag_aus_nv].should                     == 0.00
    z22[:akt_gehaltsabr_beitrag_aus_vl_gesamt].should              == 0.00
    z22[:akt_gehaltsabr_beitrag_aus_an_vl].should                  == 0.00
    z22[:akt_gehaltsabr_gesamt_brutto].should                      == 2040.00
    z22[:akt_gehaltsabr_steuern].should                            == 256.08
    z22[:akt_gehaltsabr_sv_beitraege].should                       == 0 #231.54
    z22[:akt_gehaltsabr_netto_gehalt].should                       == 1783.92
    z22[:akt_gehaltsabr_ueberweisung_vl].should                    == 40.00
    z22[:akt_gehaltsabr_ueberweisung_netto].should                 == 1743.92
  end

  it "sollte Zeile 29 korrekt einlesen" do
    z29 = @el.zeile(29)

    z29[:name].should                                              == "Hans Meier"
    z29[:kirchensteuer].should                                     == false
    z29[:berufsgruppe].should                                      == "sozialversicherungsfreier GGF"
    z29[:durchfuehrungsweg].should                                 == "Direktversicherung"
    z29[:verzicht_betrag].should                                   == 22.0
    z29[:akt_gehaltsabr_monatl_brutto_gehalt].should               == 10000.00
    z29[:akt_gehaltsabr_ag_anteil_vl].should                       == 15.00
    z29[:akt_gehaltsabr_beitrag_aus_nv].should                     == 0.00
    z29[:akt_gehaltsabr_beitrag_aus_vl_gesamt].should              == 0.00
    z29[:akt_gehaltsabr_beitrag_aus_an_vl].should                  == 0.00
    z29[:akt_gehaltsabr_gesamt_brutto].should                      == 10015.00
    z29[:akt_gehaltsabr_steuern].should                            == 4125.84
    z29[:akt_gehaltsabr_sv_beitraege].should                       == 0 #612.90
    z29[:akt_gehaltsabr_netto_gehalt].should                       == 5889.16
    z29[:akt_gehaltsabr_ueberweisung_vl].should                    == 40.00
    z29[:akt_gehaltsabr_ueberweisung_netto].should                 == 5849.16
  end

  it "sollte Zeile 21 korrekt einlesen" do
    z21 = @el.zeile(21)
    z21[:name].should                                              == "Gerda Müller"
    z21[:berufsgruppe].should                                       == "Angestellte/Arbeiter"
    z21[:akt_gehaltsabr_monatl_brutto_gehalt].should               == 2000.05
    z21[:akt_gehaltsabr_ag_anteil_vl].should                       == 40.00
    z21[:akt_gehaltsabr_beitrag_aus_nv].should                     == 0.00
    z21[:akt_gehaltsabr_beitrag_aus_vl_gesamt].should              == 0.00
    z21[:akt_gehaltsabr_beitrag_aus_an_vl].should                  == 0.00
    z21[:akt_gehaltsabr_gesamt_brutto].should                      == 2040.05
    z21[:akt_gehaltsabr_steuern].should                            == 283.39
    z21[:akt_gehaltsabr_sv_beitraege].should                       == 418.71
    z21[:akt_gehaltsabr_netto_gehalt].should                       == 1337.95
    z21[:akt_gehaltsabr_ueberweisung_vl].should                    == 40
    z21[:akt_gehaltsabr_ueberweisung_netto].should                 == 1297.95

    z21[:nv_monatl_brutto_gehalt].should               == 2000.05
    z21[:nv_ag_anteil_vl].should                       == 40.00
   # z21[:nv_beitrag_aus_nv].should                     == 116.32
    z21[:nv_beitrag_aus_vl_gesamt].should              == 0.00
    z21[:nv_beitrag_aus_an_vl].should                  == 0.00
   # z21[:nv_gesamt_brutto].should                      == 1923.73
   # z21[:nv_steuern].should                            == 248.50
#    z21[:nv_sv_beitraege].should                       == 394.85
#    z21[:nv_netto_gehalt].should                       == 1280.38
#    z21[:nv_ueberweisung_vl].should                    == 40
#    z21[:nv_ueberweisung_netto].should                 == 1240.38
#    z21[:nv_nettoverzicht].should                      == 57.57

    z21[:an_beitrag].should                                        == 196.23
#    z21[:ag_zuschuss].should                                       == 19.62
    z21[:gesamtbeitrag].should                                     == 215.85
  end

  it "sollte ausreichend Spalten enthalten" do
    @el.zeile(21).size.should >= SPALTEN_UEBERSCHRIFTEN.size
    @el.zeile(22).size.should >= SPALTEN_UEBERSCHRIFTEN.size
    @el.zeile(23).size.should >= SPALTEN_UEBERSCHRIFTEN.size
  end

  it "sollte Überschriften korrekt unterscheiden" do
    s23 = @el.zeile(23)
    s23[:bland_wohnsitz].should   == "Hessen"
    s23[:bland_arbeit].should     == "Niedersachsen"
  end
  #=end
  it "sollte auch die globalen Werte einlesen" do
    nf = @el.zeile(23)
    nf[:minijob_ok].should                    == false
    nf[:durchfuehrungsweg].should             == "Direktversicherung"
    nf[:verzicht_als_netto].should            == true
    nf[:vl_als_beitrag].should                == true
    nf[:ag_zuschuss].should                   == 10
    nf[:ag_zuschuss_als_absolut].should       == false
  end
end