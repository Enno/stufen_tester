
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
  :akt_gehaltsabr_ueberweisung_netto   => /Überweisung\s(Netto-Gehalt)/,

  :nv_monatl_brutto_gehalt => /monatliches \sBruttogehalt/,
  :nv_ag_anteil_vl         => /AG-Anteil VL/,
  :nv_beitrag_aus_nv       => /Beitrag aus \sNettoverzicht/,
  :nv_beitrag_aus_vl_gesamt => /Beitrag aus VL \s(gesamt)/,
  :nv_beitrag_aus_an_vl    => /Beitrag aus \sArbeitnehmeranteil VL/,
  :nv_gesamt_brutto        => /Gesamt-Brutto/,
  :nv_steuern              => /Steuern \s(inkl. Soli; Ki.-St.)/,
  :nv_sv_beitraege         => /SV-Beiträge/,
  :nv_netto_gehalt         => /Nettogehalt/,
  :nv_ueberweisung_vl      => /Überweisung VL/,
  :nv_ueberweisung_netto   => /Überweisung\s(Netto-Gehalt)/,
  :nv_netto_verzicht       => /Nettoverzicht\s(ursprüngliches Netto-aktuelles Netto)/,

  :vermoegensw_leistungen_monatl_brutto_gehalt => /monatliches \sBruttogehalt/,
  :vermoegensw_leistungen_ag_anteil_vl         => /AG-Anteil VL/,
  :vermoegensw_leistungen_beitrag_aus_nv       => /Beitrag aus \sNettoverzicht/,
  :vermoegensw_leistungen_beitrag_aus_vl_gesamt => /Beitrag aus VL \s(gesamt)/,
  :vermoegensw_leistungen_beitrag_aus_an_vl    => /Beitrag aus \sArbeitnehmeranteil VL/,
  :vermoegensw_leistungen_gesamt_brutto        => /Gesamt-Brutto/,
  :vermoegensw_leistungen_steuern              => /Steuern \s(inkl. Soli; Ki.-St.)/,
  :vermoegensw_leistungen_sv_beitraege         => /SV-Beiträge/,
  :vermoegensw_leistungen_netto_gehalt         => /Nettogehalt/,
  :vermoegensw_leistungen_ueberweisung_vl      => /Überweisung VL/,

  :an_beitrag    => /Arbeitnehmer-\sbeitrag/,
  :ag_zuschuss   => /Arbeitgeber-\szuschuss/,
  :gesamtbeitrag => /Gesamt-\sbeitrag/
}

#SPALTEN_UEBERSCHRIFTEN = SP_UEBERSCHRIFTEN_EINGABE

BEREICHE_INTERN_ZU_EXCEL = {
  :persoenliche_daten => ".",
  :aktuelle_gehaltsabrechnung => "akt"
}

SPALTEN_UEBERSCHRIFTEN_TEST = {
  :persoenliche_daten => {
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
    },
  :aktuelle_gehaltsabrechnung => {
      :monatl_brutto_gehalt => /monatliches \sBruttogehalt/,
      :ag_anteil_vl         => /AG-Anteil VL/,
      :beitrag_aus_nv       => /Beitrag aus \sNettoverzicht/,
      :beitrag_aus_vl_gesamt => /Beitrag aus VL \s(gesamt)/,
      :beitrag_aus_an_vl    => /Beitrag aus \sArbeitnehmeranteil VL/,
      :gesamt_brutto        => /Gesamt-Brutto/,
      :steuern              => /Steuern \s(inkl. Soli; Ki.-St.)/,
      :sv_beitraege         => /SV-Beiträge/,
      :netto_gehalt         => /Nettogehalt/,
      :ueberweisung_vl      => /Überweisung VL/,
      :ueberweisung_netto   => /Überweisung\s(Netto-Gehalt)/,
    },
  :netto_verzicht => {
      :monatl_brutto_gehalt => /monatliches \sBruttogehalt/,
      :ag_anteil_vl         => /AG-Anteil VL/,
      :beitrag_aus_nv       => /Beitrag aus \sNettoverzicht/,
      :beitrag_aus_vl_gesamt => /Beitrag aus VL \s(gesamt)/,
      :beitrag_aus_an_vl    => /Beitrag aus \sArbeitnehmeranteil VL/,
      :gesamt_brutto        => /Gesamt-Brutto/,
      :steuern              => /Steuern \s(inkl. Soli; Ki.-St.)/,
      :beitraege            => /SV-Beiträge/,
      :netto_gehalt         => /Nettogehalt/,
      :ueberweisung_vl      => /Überweisung VL/,
      :ueberweisung_netto   => /Überweisung\s(Netto-Gehalt)/,
      :nv_netto_verzicht    => /Nettoverzicht\s(ursprüngliches Netto-aktuelles Netto)/,
    },
  :vermoegenswirksame_leistungen => {
      :monatl_brutto_gehalt => /monatliches \sBruttogehalt/,
      :ag_anteil_vl         => /AG-Anteil VL/,
      :beitrag_aus_nv       => /Beitrag aus \sNettoverzicht/,
      :beitrag_aus_vl_gesamt => /Beitrag aus VL \s(gesamt)/,
      :beitrag_aus_an_vl    => /Beitrag aus \sArbeitnehmeranteil VL/,
      :gesamt_brutto        => /Gesamt-Brutto/,
      :steuern              => /Steuern \s(inkl. Soli; Ki.-St.)/,
      :sv_beitraege         => /SV-Beiträge/,
      :netto_gehalt         => /Nettogehalt/,
      :ueberweisung_vl      => /Überweisung VL/,
      :ueberweisung_netto   => /Überweisung\s(Netto-Gehalt)/
    },
  :arbeitgeberzuschuss => {
    :an_beitrag     => /Arbeitnehmer-\sbeitrag/,
    :ag_zuschuss    => /Arbeitgeber-\szuschuss/,
    :gesamtbeitrag  => /Gesamt-\sbeitrag/
  }
}

#ABSCHNITTS_UEBERSCHRIFTEN = {
#  :persoenliche_daten            => nil,
#  :aktuelle_gehaltsabrechnung    => /Aktuelle Gehaltsabrechnung/,
#  :netto_verzicht                => /Nettoverzicht x €/,
#  :vermoegenswirksame_leistungen => /vermögenswirksame Leistungen/,
#  :arbeitgeberzuschuss           => nil
#}

EXCEL_EINLESE_TRANSFORMATIONEN = {
  :minijob_ok               => {"ja"     => true, "j"     => true, "nein"   => false, "n"   => false},
  :kirchensteuer            => {"ja"     => true, "j"     => true, "nein"   => false, "n"   => false},
  :vl_als_beitrag           => {"ja"     => true, "j"     => true, "nein"   => false, "n"   => false},
  :kinderlos                => {"ja"     => true, "j"     => true, "nein"   => false, "n"   => false},
  :pausch_steuer40b         => {"ja"     => true, "j"     => true, "nein"   => false, "n"   => false},
  :ag_zuschuss_als_absolut  => {"€"     => true, "%"      => false},
  :verzicht_als_netto       => {"netto" => true, "brutto" => false},
  :berufsgruppe             => {"sozialvers.freier GGF" => "sozialversicherungsfreier GGF",
    "Angestellte/Arbeiter"  => "Angestellte/Arbeiter",
    "Azubi"                 => "Azubi"}
}