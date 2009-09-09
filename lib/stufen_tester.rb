
require 'win32ole'

require 'grundlage'
require 'form_filler'
require 'excel_leser'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class StufenTester
  def initialize(source_path, source_file, destination_file_path, destination_file, start_proc_name)
    @source_data = {}
    @source_path = source_path
    @source_file = source_file
    @destination_file_path = destination_file_path
    @destination_file = destination_file
    @start_proc_name = start_proc_name
  end
  def open_source_file
    @el = ExcelLeser.new(@source_path + @source_file, "Global", "Tabelle")
  end
  def open_destination_file
    @ff = FormFiller.new(@destination_file_path, @destination_file, @start_proc_name)
    2.times { @ff.send_keys("%{F11}") } # Workaround für nicht funktionierende Tab-Tasten
  end

  def readin_source_data(row)
    open_source_file
    @source_data = @el.zeile(row)
  end

  def write_source_data_into_template(source_data)
    open_destination_file
    @ff.populate_template(source_data)
  end

  def call_destination_function(vb_function_name, *args)
    @ff.vb_send(vb_function_name, *args)
  end

  def close_source_file
    @el.excel_beenden
  end

  def close_destination_file
    @ff.quit_excel
  end
  
end
