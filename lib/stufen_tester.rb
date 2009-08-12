
#ORIG_ANORDNUNG = [
#  :name,
#  :bruttogehalt,
#  :freibetrag,
#
#]
#
#zeile[:freibetrag] =
require 'win32ole'

require 'grundlage'
require 'form_filler'
require 'excel_leser'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class StufenTester
  def initialize
    @source_data = {}
  end
  def open_source_file
    source_file = "test.xls"
    source_path = "C:/Praktikum/stufen_tester" +  "/daten/" #{}File.dirname(File.dirname(__FILE__)) +  "/daten/"
    puts source_path
    @el = ExcelLeser.new(source_path + source_file, "Global", "Tabelle")
  end
  def open_destination_file
    destination_file = "sr38a_entkernt_test.xls"
    destination_file_path = "C:/Praktikum/stufen_tester" +  "/daten/"  #File.dirname(File.dirname(__FILE__)) +  "/daten/"
    start_proc_name = "Entgeltumwandlungsrechner_starten"
    @ff = FormFiller.new(destination_file_path, destination_file, start_proc_name)
    2.times { @ff.send_keys("%{F11}") } # Workaround für nicht funktionierende Tab-Tasten
  end

  def readin_source_data(row)
    open_source_file
    @source_data = @el.zeile(row)
    return @source_data
  end

  def write_source_data_into_template
    open_destination_file
    @ff.populate_template(@source_data)
  end

  def close_source_file
    @el.excel_beenden
  end

  def close_destination_file
    @ff.quit_excel
  end
  
end
