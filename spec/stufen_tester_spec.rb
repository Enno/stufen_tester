puts "stufen_tester_spec"
require 'lib/stufen_tester'

describe StufenTester do
  before(:each) do
    source_file = "test.xls"
    source_path = File.dirname(File.dirname(__FILE__)) +  "\\daten\\"
    destination_file = "sr38a_op_tor2.xls"
    destination_file_path = File.dirname(File.dirname(__FILE__)) +  "\\daten\\"
    start_proc_name = "Entgeltumwandlungsrechner_starten"
    @stufen_tester = StufenTester.new(source_path, source_file, destination_file_path, destination_file, start_proc_name)
  end
  after(:each) do
  end

  #  def sende_tasten(*args, &blk)
  #    @tasten_sender.sende_tasten(*args, &blk)
  #  end
  #
  #  it "should desc" do
  #    stuf_rech_pfad = "H:/GiS/gm/gMisc/VertriebStufen/StufenR/stufenrechner_Version_3_8a_offen_passiviert.xls"
  #    #system "start excel #{stuf_rech_pfad}"
  #    excel1 = WIN32OLE.new('Excel.Application')
  #    excel1.Visible = true
  #    strechner = excel1.Workbooks.Open(stuf_rech_pfad)
  #    sende_tasten('Microsoft Excel', nil).should == true
  #    sende_tasten('Microsoft Excel - stufenrechner', nil).should == true
  #
  #  end


  #  it "sollte existieren" do
  #    @stufen_tester.should_not be_nil
  #  end
  #
  #  it "sollte excel (source) oeffnen" do
  #    @stufen_tester.open_source_file
  #    @stufen_tester.close_source_file
  #  end
  #
  #  it "sollte excel (destination) oeffnen" do
  #    @stufen_tester.open_destination_file
  #    @stufen_tester.close_destination_file
  #  end

  #  it "sollte zeile 22 einlesen" do
  #    z22 = @stufen_tester.readin_source_data(22)
  #    z22[:name].should                  == "Hans Meier"
  #    z22[:verzicht_betrag].should       == 50.0
  #    @stufen_tester.close_source_file
  #  end
  
  it "sollte alle zeilen einlesen und ins template einfuegen" do
    #[0,1,nil,3,4,5,6,7,nil].each do |i|
    #[2,8].each do |i| #TODO: fehler beheben
    [0].each do |i|
      next unless i
      
      zeilennr = 21 + i
      zeile = @stufen_tester.readin_source_data(zeilennr)
      puts zeile.inspect

      @stufen_tester.write_source_data_into_template(zeile)
      keys_zu_stufenrechner_namen = {
        :name                    => "name",
        :bruttogehalt            => "gehalt",
        #:freibetrag              => "Freibetrag",
        #"kv_pflicht",
        #"KV_privat",
        #    :steuerklasse       => "Steuerklasse",
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
        [i, key, @stufen_tester.call_destination_function("Abfrage_Feld_#{sr_name}")].should == [i, key, zeile[key]]
      end
      puts "Zeile: #{i}: #{keys_zu_stufenrechner_namen.size} Felder getestet"
      #@stufen_tester.call_destination_function("Abfrage_Ergebnis", "monatlichesbruttogehalt", "akt").should == zeile[:akt_gehaltsabr_monatl_brutto_gehalt]
      #@stufen_tester.call_destination_function("Abfrage_Ergebnis", "aganteilvl", "akt").should              == zeile[:akt_gehaltsabr_ag_anteil_vl  ]
      @stufen_tester.call_destination_function("Abfrage_Ergebnis", "beitragausnettoverzicht", "akt").should  == zeile[:akt_gehaltsabr_beitrag_aus_nv]
      @stufen_tester.call_destination_function("Abfrage_Ergebnis", "beitragausvl", "akt").should             == zeile[:akt_gehaltsabr_beitrag_aus_vl_gesamt]
      @stufen_tester.call_destination_function("Abfrage_Ergebnis", "beitragausananteilvl", "akt").should     == zeile[:akt_gehaltsabr_beitrag_aus_an_vl]
      #@stufen_tester.call_destination_function("Abfrage_Ergebnis", "gesamtbrutto", "akt").should             == zeile[:akt_gehaltsabr_gesamt_brutto]
      @stufen_tester.call_destination_function("Abfrage_Ergebnis", "steuern", "akt").should                  == zeile[:akt_gehaltsabr_steuern]
      @stufen_tester.call_destination_function("Abfrage_Ergebnis", "svbeiträge", "akt").should               == zeile[:akt_gehaltsabr_sv_beitraege]
      @stufen_tester.call_destination_function("Abfrage_Ergebnis", "überweisungvl", "akt").should            == zeile[:akt_gehaltsabr_ueberweisung_vl]
      #@stufen_tester.call_destination_function("Abfrage_Ergebnis", "überweisung", "akt").should              == zeile[:akt_gehaltsabr_ueberweisung_netto]
      @stufen_tester.close_source_file
      @stufen_tester.close_destination_file
    end
  end

end


