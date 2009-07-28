# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'lib/stufen_tester'
require 'lib/tasten_sender'

describe StufenTester do
  before(:each) do
    @stufen_tester = StufenTester.new
    @tasten_sender = TastenSender.new #(:wartezeit => 0.3)
  end

  def sende_tasten(*args, &blk)
    @tasten_sender.sende_tasten(*args, &blk)
  end

  it "should desc" do
    stuf_rech_pfad = "H:/GiS/gm/gMisc/VertriebStufen/StufenR/stufenrechner_Version_3_8a_offen_passiviert.xls"
    #system "start excel #{stuf_rech_pfad}"
    excel1 = WIN32OLE.new('Excel.Application')
    excel1.Visible = true
    strechner = excel1.Workbooks.Open(stuf_rech_pfad)
    sende_tasten('Microsoft Excel', nil).should == true
    sende_tasten('Microsoft Excel - stufenrechner', nil).should == true
    
    
    # TODO
  end
end

