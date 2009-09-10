
GLOBAL_UEBERSCHRIFTEN = {
  # blatt 1 (Global) noch ausdokumentiert
  #  :minijob_ok       => "Berechnung ggf. auch im  Minijob-Bereich darstellen",
  #  :durchfuehrungsweg=> "Durchführungsweg",
  #  :verzicht_als_netto=> "Betrag ist der Netto-/ Brutto-Verzicht",
  #  :vl_als_beitrag   => "Verwendung der VL als Beitrag",
  #  :ag_zuschuss      => "Angabe in €/ in % des Umwandlungsbetrages",
  #  :ag_zuschuss_als_absolut=> "Prozent / absolut",
}

GLOBALBLATT_NAMEN = {
  :minijob_ok               => "MinijobOK",
  :durchfuehrungsweg        => "Durchführungsweg",
  :verzicht_als_netto       => "NettoOderBrutto",
  :vl_als_beitrag           => "VLAlsBeitragVerwenden",
  :ag_zuschuss              => "ArbeitgeberZuschuss",
  :ag_zuschuss_als_absolut  => "AGZuschussProzentOderAbsolut",
  
}

SPALTEN_UEBERSCHRIFTEN = {
  :name             => /Name, Vorname/,
	:personal_nr      => /Personalnr./,
	:geb_datum        => /Geburtsdatum/,
	:geschlecht       => /Geschlecht/,
	:bruttogehalt     => /Bruttogehalt (mtl.)/	,
  :freibetrag       => /Freibetrag/,
  :k_vers_art       => /Kranken-\sversicherung/,
	:steuerklasse     => /Steuer-\sklasse/,
  :kirchensteuer    => /Kirchen-\ssteuer/,
  :kinder_fb        => /Kinder-\sfreibetrag/,
  :kinderlos        => /Kinderlos\s(erhöhter PV-Satz für unter 23-jährige)/,
  :verzicht_betrag  => /Netto-\/\s*Bruttoverzicht/,
  :vl_arbeitgeber   => /Arbeitgeber-\santeil VL/,
  :vl_arbeitnehmer  => /Arbeitnehmer-\santeil VL/,
  #:vl_gesamt        => nil,
  :bland_wohnsitz   => /Bundesland \sWohnsitz/,
  :bland_arbeit     => /Bundesland\sArbeitsstätte/,
  :berufsgruppe     => /Berufsgruppe/,
  :pausch_steuer40b	=> /Pauschalversteuerung Nach \s40b EStG wird aktuell genutzt/,

  :akt_gehaltsabr_monatl_brutto_gehalt => /monatliches \sBruttogehalt/,
  :akt_gehaltsabr_ag_anteil_vl         => /AG-Anteil VL/,
  :akt_gehaltsabr_beitrag_aus_nv       => /Beitrag aus \sNettoverzicht/,
  :akt_gehaltsabr_beitrag_aus_vl_gesamt => /Beitrag aus VL \s(gesamt)/,
  :akt_gehaltsabr_beitrag_aus_an_vl    => /Beitrag aus \sArbeitnehmeranteil VL/,
  :akt_gehaltsabr_gesamt_brutto        => /Gesamt-Brutto/,
  :akt_gehaltsabr_steuern              => /Steuern \s(inkl. Soli; Ki.-St.)/,
  :akt_gehaltsabr_sv_beitraege         => /SV-Beiträge/,
  :akt_gehaltsabr_netto_gehalt         => /Nettogehalt/,
  :akt_gehaltsabr_ueberweisung_vl      => /Überweisung VL/,
  :akt_gehaltsabr_ueberweisung_netto   => /Überweisung\s(Netto-Gehalt)/
}

#ABSCHNITTS_UEBERSCHRIFTEN = {
#  :persoenliche_daten            => nil,
#  :aktuelle_gehaltsabrechnung    => /Aktuelle Gehaltsabrechnung/,
#  :netto_verzicht                => /Nettoverzicht x €/,
#  :vermoegenswirksame_leistungen => /vermögenswirksame Leistungen/,
#  :arbeitgeberzuschuss           => nil
#}

EXCEL_EINLESE_TRANSFORMATIONEN = {
  :minijob_ok               => {"ja"    => true, "nein"   => false},
  :kirchensteuer            => {"j"     => true, "n"      => false},
  :vl_als_beitrag           => {"ja"    => true, "nein"   => false},
  :kinderlos                => {"j"     => true, "n"      => false},
  :pausch_steuer40b         => {"j"     => true, "nein"   => false},
  :ag_zuschuss_als_absolut  => {"€"     => true, "%"      => false},
  :verzicht_als_netto       => {"netto" => true, "brutto" => false},
  :berufsgruppe             => {"sozialvers.freier GGF" => "sozialversicherungsfreier GGF",
    "Angestellte/Arbeiter"  => "Angestellte/Arbeiter",
    "Azubi"                 => "Azubi"}
}