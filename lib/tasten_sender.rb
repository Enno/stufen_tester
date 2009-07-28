
require 'win32ole'

class TastenSender
  def initialize(optionen={})
    @wsh = WIN32OLE.new('Wscript.Shell')
    @wartezeit = optionen[:wartezeit] || 0.2
  end

  def sende_tasten(fenstername, tastenfolge, optionen={})
    wartezeit = optionen[:wartezeit] || @wartezeit
    sleep wartezeit
    fenster_aktiv = fenstername ? @wsh.AppActivate(fenstername) : true
    if fenster_aktiv
      @wsh.SendKeys(tastenfolge) if tastenfolge
      yield if block_given?
    else
      else_aktion = optionen[:fenster_fehlt]
      case else_aktion
      when Proc
        else_aktion.call
      when nil
      else
        raise else_aktion
      end
    end
    fenster_aktiv
  end
end
