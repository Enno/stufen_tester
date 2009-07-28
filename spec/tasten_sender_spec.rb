# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'lib/tasten_sender'

describe TastenSender do
  before(:each) do
    @ts = TastenSender.new(:wartezeit => 1.8)
  end

  def sende_tasten(*args, &blk)
    @ts.sende_tasten(*args, &blk)
  end
  
  it "should be able to control notepad" do
    @ts.should_not be_nil

    sende_tasten('Editor', '%{F4}') # anderen Editor schließen, falls keine ungespeicherten Änderungen
    sende_tasten('Editor', nil).should == false # "Anderer Editor vorher schon geöffnet"
    
    system "start notepad"
    sende_tasten('Editor', nil).should == true

    sende_tasten('Editor', 'Ruby{TAB}on{TAB}Windows{ENTER}', :fenster_fehlt => "Editor nicht gestartet") do
      # ALT-F to pull down File menu, then A to select Save As...:
      sende_tasten(nil, '%DU')
      sende_tasten('Speichern unter', 'H:\temp_filename.txt{ENTER}') do
        sende_tasten('Speichern unter', 'J')
      end
      # Quit Notepad with ALT-F4:
      sende_tasten('Editor', '%{F4}', :fenster_fehlt => "kein Editor mehr gefunden")
    end
    
    sende_tasten('Editor', nil).should == false
    
  end
end

