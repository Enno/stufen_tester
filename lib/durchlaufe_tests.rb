
require 'win32ole'

# Create an instance of the Wscript Shell:
wsh = WIN32OLE.new('Wscript.Shell')
# Try to activate the Notepad window:
if false and gefunden=wsh.AppActivate('Editor')
     sleep(1)
    # Enter text into Notepad:
    wsh.SendKeys('Ruby{TAB}on{TAB}Windows{ENTER}')
    # ALT-F to pull down File menu, then A to select Save As...:
    wsh.SendKeys('%DU')
    #wsh.SendKeys('U')
    sleep(1)
    if wsh.AppActivate('Speichern unter') #('Save As')
        wsh.SendKeys('H:\temp_filename.txt{ENTER}')
        sleep(1)
        # If prompted to overwrite existing file:
        if wsh.AppActivate('Speichern unter')
            # Enter 'Y':
            wsh.SendKeys('J')
        end
    end
    # Quit Notepad with ALT-F4:
    wsh.SendKeys('%{F4}')
end
#puts "Hello World"
#puts "nicht gefunden" unless gefunden

$wsh = WIN32OLE.new('Wscript.Shell')
module TastenSender
  def sende_tasten(fenstername, tastenfolge, optionen={})
    wartezeit = 0.1 || optionen[:wartezeit]
    sleep wartezeit
    fenster_aktiv = fenstername ? $wsh.AppActivate(fenstername) : true
    if fenster_aktiv 
      $wsh.SendKeys(tastenfolge) if tastenfolge
      yield if block_given?
    else
      else_proc = optionen[:falls_fenster_nicht_da]
      else_proc.call if else_proc
    end
  end
end

include TastenSender
# Enter text into Notepad:
sende_tasten('Editor', 'Ruby{TAB}on{TAB}Windows{ENTER}') do
    # ALT-F to pull down File menu, then A to select Save As...:
    sende_tasten(nil, '%DU')
    sende_tasten('Speichern unter', 'H:\temp_filename.txt{ENTER}') do
        sende_tasten('Speichern unter', 'J')
    end
    # Quit Notepad with ALT-F4:
    wsh.SendKeys('%{F4}')
end

