puts "stufen_tester_spec"
require 'lib/stufen_tester'

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

keys_zu_vb_abfrage_namen = {
  :akt_gehaltsabr_monatl_brutto_gehalt  => "monatlichesbruttogehalt",
  :akt_gehaltsabr_ag_anteil_vl          => "aganteilvl",
  :akt_gehaltsabr_beitrag_aus_nv        => "beitragausnettoverzicht",
  :akt_gehaltsabr_beitrag_aus_vl_gesamt => "beitragausvl",
  :akt_gehaltsabr_beitrag_aus_an_vl     => "beitragausananteilvl",
  :akt_gehaltsabr_gesamt_brutto         => "gesamtbrutto",
  :akt_gehaltsabr_steuern               => "steuern",
  :akt_gehaltsabr_sv_beitraege          => "svbeiträge",
  :akt_gehaltsabr_ueberweisung_vl       => "überweisungvl",
  :akt_gehaltsabr_ueberweisung_netto    => "überweisung"
}

  
describe StufenTester do
  before(:all) do
    source_file = "test_more.xls"
    source_path = File.dirname(File.dirname(__FILE__)) +  "\\daten\\"
    destination_file = "sr38a_op_tor2.xls"
    destination_file_path = File.dirname(File.dirname(__FILE__)) +  "\\daten\\"
    start_proc_name = "Entgeltumwandlungsrechner_starten"
    @stufen_tester = StufenTester.new(source_path, source_file, destination_file_path, destination_file, start_proc_name)
  end

  #[1, -2, -3, -4, -5, -6, -7, -8, -9]
  # Probl: 11, 12, 13
  [-10, 14].each do |i|
    next if i.nil? or i < 0
  
    describe StufenTester, "in Zeile #{i}" do
      before(:all) do
        zeilennr = 20 + i
        @zeile = @stufen_tester.readin_source_data(zeilennr)
        puts @zeile.inspect
        @stufen_tester.write_source_data_into_template(@zeile)
      end

      after(:all) do
        @stufen_tester.close
      end
   
      keys_zu_stufenrechner_namen.each do |key, sr_name|
        it "sollte bei #{key} mit Stufenrechner-Feld #{sr_name} übereinstimmen" do
          @stufen_tester.check_reference_data("Abfrage_Feld_#{sr_name}").should == @zeile[key]
        end
      end

      keys_zu_vb_abfrage_namen.each do |key, vb_name|
        it "sollte bei #{key} mit VB-Abfrage-Feld #{vb_name} übereinstimmen" do
          @zeile[key].should == @stufen_tester.check_reference_data("Abfrage_Ergebnis", vb_name, "akt")
        end
      end

    end
  end
end


