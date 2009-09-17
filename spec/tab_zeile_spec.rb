# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'tab_zeile'

describe TabZeile do
  before(:each) do
    @tz = TabZeile.new
#    @tz.bereich(".")[:name] = "Heinz"
    @tz[:name, "."] = "Heinz"
  end

  it "sollte werte abrufen können" do
    @tz[:name,"."].should == "Heinz"
  end

  it "sollte eingaben auslesen" do
    @tz.eingaben[:name].should == "Heinz"
  end

  it "sollte eingaben zuweisen können" do
    @tz.eingaben[:name] = "Hans"
    @tz.eingaben[:name].should == "Hans"
    @tz[:name].should == "Hans"
  end

end

