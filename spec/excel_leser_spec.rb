# enconding: utf8

require File.dirname(File.dirname(__FILE__)) +  '/lib/excel_leser.rb'


describe ExcelLeser do
  before(:each) do
    mappen_name = "test.xls"
    mappen_pfad = File.dirname(File.dirname(__FILE__)) +  "/daten/"
    @el = ExcelLeser.new(mappen_pfad + mappen_name, "Global", "Tabelle")
  end
  after(:each) do
    @el.excel_beenden
  end

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

    z22[:name].should                  == "Hans Meier"
    z22[:verzicht_betrag].should       == 50.0
  end

  it "sollte Zeile 21 korrekt einlesen" do
    z21 = @el.zeile(21)
    # z21[:name][0..6].should                  == "Gerda M" #üller"
    z21[:name].should                  == "Gerda Müller"
    z21[:berufsgruppe].should          == "Angestellte/Arbeiter"
  end

  it "sollte ausreichend Spalten enthalten" do
    @el.zeile(21).size.should >= SPALTEN_UEBERSCHRIFTEN.size
    @el.zeile(22).size.should >= SPALTEN_UEBERSCHRIFTEN.size
    @el.zeile(23).size.should >= SPALTEN_UEBERSCHRIFTEN.size
  end

  it "sollte Überschriften korrekt unterscheiden" do
    s23 = @el.zeile(23)
    s23[:bland_wohnsitz].should == "Hessen"
    s23[:bland_arbeit].should == "Niedersachsen"
  end

  it "sollte auch die globalen Werte einlesen" do
    z21 = @el.zeile(21)
    z21[:minijob_ok].should == "nein"
  end

  #it "sollte nur Überschriften, die im Blatt existieren, finden" do
    #z23 = @el.zeile(23)
    #z23[:berufsgruppe].should == "Angestellte/Arbeiter"
    #proc{ z23[:berufsgruppe] }.should   raise_error
    #proc{ s23[:kirchensteuer] }.should   raise_error
  #end

=begin
  it "sollte auf die Methode 'spalte' reagieren und einen Hash zurückgeben" do
    @el.spalte("Name, Vorname").is_a?(Hash).should be_true
  end

  it "sollte Spalte 'name' korrekt einlesen" do
    s1 = @el.spalte("Name, Vorname")
    s1[0].should                  == "Gerda Müller"
    s1[1].should                  == "Hans Meier"
  end

  it "sollte Spalte 'verzicht_betrag' korrekt einlesen" do
    s1 = @el.spalte("Netto-/Bruttoverzicht")
    s1[0].should                  == 57.57
    s1[1].should                  == 50.00
  end
=end
end

