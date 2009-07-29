
#constant definition
FILE_PATH = File.dirname(File.dirname(__FILE__)) +  '/daten/'

class ExcelInputBox # spaeter ersetzen durch rwd http://www.erikveen.dds.nl/rubywebdialogs/index.html#1.0.0
  require 'win32ole'
  def initialize
    @user_iput = ""
  end
  def get_user_input(prompt='')
    #@user_input = ""
    until @user_input.is_a?(String) && @user_input.length > 0 do
      excel_input_box = WIN32OLE.new('Excel.Application')
      @user_input = excel_input_box.InputBox(prompt)
      excel_input_box.Quit
      excel_input_box = nil
    end
    return @user_input
  end
end
class ExcelController
  require 'win32ole'
  attr_accessor :excel_file_name, :excel_sheet_name
  attr_reader :workbook, :sheet
  def initialize()
    @excel_file_name = ""
    @excel_appl = ""
    @workbook = ""
    @sheet = Array.new
    sheet_count = 0
  end
  def open_excel_file()
    if File.exist?(@excel_file_name)
      @excel_appl = WIN32OLE.new('Excel.Application')
      @excel_appl['Visible'] = true
      @excel_appl.Workbooks.Open(FILE_PATH + @excel_file_name)
      @workbook = WIN32OLE.connect('Excel.Application').ActiveWorkbook #problem wenn alter excel task noch aktiv
    else
      raise "Error Message: " "Datei nicht vorhanden" # eventuell neue abfrage
    end
  end
  def find_excel_sheet
    sheet_count = @excel_appl.Worksheets.count
    for i in (1..sheet_count) do #@excel_appl.Worksheets.each) do
      @sheet << @excel_appl.Worksheets(i).Name
    end
    return @sheet
  end
  def select_excel_sheet
    if @sheet.include?(@excel_sheet_name)
    puts @excel_sheet_name
    @excel_appl.Worksheets(@excel_sheet_name).select
    else
      raise "Error Message: " "Sheet nicht vorhanden" # eventuell neue abfrage
    end
  end
  def close_excel_file()
    @excel_appl.ActiveWorkbook.Close(0)
  end
  def quit_excel
    @excel_appl.quit
    @excel_appl = nil
    GC.start #garbage control starten
  end
end
class ExcelReader
  require 'win32ole'
  def initialize()

  end
  def get_contents
    
  end

end


excel = ExcelController.new

excel.excel_file_name = ExcelInputBox.new.
  get_user_input("Welche Datei soll geoeffnet werden?") + '.xls'

excel.open_excel_file

excel.excel_sheet_name = ExcelInputBox.new.
  get_user_input("Welches Worksheet soll geoeffnet werden? (" +
    excel.find_excel_sheet.collect { |names| names + ", "}.to_s.chop.chop + ")")

excel.select_excel_sheet

sleep(2)
excel.close_excel_file
sleep(2)
excel.quit_excel