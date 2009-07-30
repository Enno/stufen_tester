
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
  attr_reader :workbook, :sheet, :current_sheet_name
  def initialize()
    @excel_file_name = ""
    @excel_sheet_name = ""
    @excel_appl = ""
    @workbook = ""
    @sheet = Array.new
    sheet_count = 0
    @@current_sheet_name = ""
    @@current_excel_appl = ""
  end
  def open_excel_file()
    if File.exist?(@excel_file_name)
      #      @excel_appl = WIN32OLE.GetActiveObject('Excel.Application') ||
      #        WIN32OLE.new('Excel.Application')
      @excel_appl = WIN32OLE.new('Excel.Application') 
      @excel_appl['Visible'] = true
      @excel_appl.Workbooks.Open(FILE_PATH + @excel_file_name)
      @@current_excel_appl = @excel_appl
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
      @excel_appl.Worksheets(@excel_sheet_name).select
      @@current_sheet_name = @excel_sheet_name
      return @@current_sheet_name
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
# person + personalnr zur identifikation, falls doppelte namen
# spaeter abfrage welchen datensatz auslesen
class ExcelReader < ExcelController
  require 'win32ole'
  require 'stufen_tester'
  attr_accessor :column_name, :first_std_column
  def initialize()
    @start_row = 1
    @start_colum = 1
    @start_address = nil
    @row_offset = 2
    @column_offset = 0
    @column_name = ""
    @one_dataset = []
    @current_value = []
    i = 0
    @first_std_column = ""
  end
  def get_contents
    WIN32OLE.connect('Excel.Application').ActiveWorkbook
    if @@current_excel_appl.
        WorkSheets(@@current_sheet_name).Cells(@start_row, @start_colum).
        find(@column_name[0..6])  #
      @start_address = @@current_excel_appl.
        WorkSheets(@@current_sheet_name).Cells(@start_row, @start_colum).
        find(@column_name).address
      @current_value = @@current_excel_appl.WorkSheets(@@current_sheet_name).
        range(@start_address).offset(@row_offset,@column_offset).value
      if @current_value == nil
        raise "Error Message: " "Kein Datensatz vorhanden."
      else
        i=0
        until @current_value == nil #bessere schleife einbauen
          @current_value = @@current_excel_appl.WorkSheets(@@current_sheet_name).
            range(@start_address).offset(@row_offset+i,@column_offset).value
          i+=1
          @one_dataset << @current_value
          puts @current_value
        end
      end
      return @one_dataset
    else
      raise "Error Message :" "Spaltenname nicht gefunden"
    end
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

datasets = ExcelReader.new

column_name = ExcelInputBox.new.
  get_user_input("Welche Spalte soll ausgelesen werden?") # pulldown menue benoetigt 
if SPALTEN_UEBERSCHRIFTEN.include?(column_name.to_sym)
  datasets.column_name = SPALTEN_UEBERSCHRIFTEN["#{column_name}".to_sym]
  datasets = datasets.get_contents
elsif column_name == "all"
  SPALTEN_UEBERSCHRIFTEN.each do |key, value|
  puts SPALTEN_UEBERSCHRIFTEN[key]
  datasets.column_name = SPALTEN_UEBERSCHRIFTEN[key]
    datasets[key] = datasets.get_contents
  end
else
  raise "Error Message :" "Spaltenname existiert nicht"
end
puts datasets.inspect


#sleep(2)
#excel.close_excel_file
#sleep(2)
#excel.quit_excel