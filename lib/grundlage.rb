
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
  :kinderlos        => /Kinderlos/,
  :verzicht_betrag  => /Netto-\/\s*Bruttoverzicht/,
  :vl_arbeitgeber   => /Arbeitgeber-\santeil VL/,
  :vl_arbeitnehmer  => /Arbeitnehmer-\santeil VL/,
  :vl_gesamt        => nil,
  :bland_wohnsitz   => /Bundesland \sWohnsitz/,
  :bland_arbeit     => /Bundesland\sArbeitsstätte/,
  :berufsgruppe     => /Berufsgruppe/,
  :pausch_steuer40b	=> /Pauschalversteuerung Nach \s40b EStG wird aktuell genutzt/
}

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