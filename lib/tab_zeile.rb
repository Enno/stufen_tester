# To change this template, choose Tools | Templates
# and open the template in the editor.

class TabZeile
  def initialize
    @bereiche = { "." => {}, "akt" => {}, "nv" => {}, "vl" => {}, "erg" => {} }
  end

  def [](key, bereich_id = ".")
    bereich(bereich_id)[key]
  end

  def []=(key, bereich_id, wert)
    bereich(bereich_id)[key] = wert
  end

  def bereich(bereich_id)
    @bereiche[bereich_id.to_s]
  end

  def eingaben
    bereich(".")
  end
end
