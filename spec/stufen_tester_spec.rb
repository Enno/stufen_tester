require 'lib/stufen_tester'

keys_zu_stufenrechner_namen = {
  :name                    => "name",
  :bruttogehalt            => "gehalt",
  :freibetrag              => "Freibetrag",
  :k_vers_art              => "kv_pflicht",
  #"KV_privat",
  :steuerklasse            => "Steuerklasse",
  :kinder_fb               => "kinderfreibetraege",
  :kirchensteuer           => "Kirchensteuer",
  :bland_wohnsitz          => "Wohnsitz",
  :bland_arbeit            => "arbeitsstaette",
  :berufsgruppe            => "Berufsgruppe",
  :durchfuehrungsweg       => "bavweg",
  :pausch_steuer40b        => "dive_40b_vorhanden",
  :minijob_ok              => "Minijob",
  :kinderlos               => "erh_pvsatz",

  #   :nvz            => "nvz",
  :verzicht_betrag         => "nvz_betrag",
  :verzicht_als_netto      => "nvz_netto",
  #:verzicht_als_netto      => "nvz_brutto",
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

keys_zu_stufenrechner_trafos = {
  :k_vers_art      => proc {|kv_art| kv_art == "g"},
  :steuerklasse    => proc {|stkl|   (%w[I II III IV V VI].index(stkl) + 1).to_s }

}

keys_zu_vb_abfrage_namen = {
  :monatl_brutto_gehalt  => "monatlichesbruttogehalt",
  :ag_anteil_vl          => "aganteilvl",
  :beitrag_aus_nv        => "beitragausnettoverzicht",
  :beitrag_aus_vl_gesamt => "beitragausvl",
  :beitrag_aus_an_vl     => "beitragausananteilvl",
  :gesamt_brutto         => "gesamtbrutto",
  :steuern               => "steuern",
  :sv_beitraege          => "svbeiträge",
  :ueberweisung_vl       => "überweisungvl",
  :ueberweisung_netto    => "überweisung"
}
describe "" do

  [3].each do |suffix|
    dateiname = "tg-#{suffix}.xls"

  
    describe "Test #{dateiname}" do
      before(:all) do
        source_path = File.dirname(File.dirname(__FILE__)) +  "\\daten\\"
        destination_file = "sr38a_op_tor2.xls"
        destination_file_path = File.dirname(File.dirname(__FILE__)) +  "\\daten\\"
        start_proc_name = "Entgeltumwandlungsrechner_starten"
        @stufen_tester = StufenTester.new(source_path, dateiname, destination_file_path, destination_file, start_proc_name)
      end

      after :all do
        @stufen_tester.close_test_table
      end


      #[1, -2, -3, -4, -5, -6, -7, -8, -9]
      # Probl: 11, 12, 13
      #[13, -14].each do |i|
      #[-1, -2, -3, -4, -5, -6, -7, -8, -9, -10, -11, -12, -13, 14].each do |i|
      #[11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21].each do |i|
      [94].each do |i|
      #[12].each do |i|
      #(10..100).each do |i|
        next if i.nil? or i < 0

        describe StufenTester, "in Zeile #{i}" do
          before(:all) do
            zeilennr = 20 + i
            @zeile = @stufen_tester.readin_source_data(zeilennr)
            @stufen_tester.write_source_data_into_template(@zeile)
          end

          after(:all) do
            @stufen_tester.close_reference_excel
          end

          def vergleiche(ist_wert, soll_wert)
            if soll_wert.is_a? Numeric
              delta = 1e-6
              ist_wert.should be_close soll_wert, delta
            else
              ist_wert.should == soll_wert
            end
          end

          keys_zu_stufenrechner_namen.each do |key, sr_name|
            it "sollte Feld #{key} in das Stufenrechner-Feld #{sr_name} übernehmen" do
              stufenr_wert = @stufen_tester.check_reference_data("Abfrage_Feld_#{sr_name}")
              tab_wert = @zeile[key]

              trafo = keys_zu_stufenrechner_trafos[key]
              tab_wert = trafo[tab_wert] if trafo
              # hack
              tab_wert = true if key == :verzicht_als_netto and @zeile[:verzicht_betrag] == 0

              vergleiche stufenr_wert, tab_wert

            end
          end

          #["akt"]
          %w[akt nv vl].each do |excel_bereich|
            keys_zu_vb_abfrage_namen.each do |key, vb_name|
              it "sollte in #{excel_bereich} bei #{key} mit VB-Abfrage-Feld #{vb_name} übereinstimmen" do
                vergleiche @zeile[key, excel_bereich], @stufen_tester.check_reference_data("Abfrage_Ergebnis", vb_name, excel_bereich)
              end
            end
          end
        end
      end
    end

  end
end