puts "stufen_tester_spec"
require 'lib/stufen_tester'

describe StufenTester do
  before(:each) do
    @stufen_tester = StufenTester.new
    #    @tasten_sender = TastenSender.new #(:wartezeit => 0.3)
  end
  after(:each) do
    @stufen_tester.close_source_file
    @stufen_tester.close_destination_file
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
  #  end
  #
  #  it "sollte excel (destination) oeffnen" do
  #    @stufen_tester.open_destination_file
  #  end

  it "sollte zeile 22 einlesen" do
    z22 = @stufen_tester.readin_source_data(22)
    z22[:name].should                  == "Hans Meier"
    z22[:verzicht_betrag].should       == 50.0
  end

  it "sollte zeile 22 einlesen und ins template einfuegen" do
    #z22 = @stufen_tester.readin_source_data(22)
    @stufen_tester.write_source_data_into_template
  end

end


