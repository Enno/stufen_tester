
#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

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
  attr_accessor :excel_datei_name, :excel_blatt_name
  attr_reader :workbook, :sheet, :current_sheet_name
  def initialize(pfad, blatt1, blatt2)
    @excel_datei_name = ""
    @excel_blatt_name = ""
    @excel_appl = ""
    @mappe = ""
    @sheet = Array.new
    sheet_count = 0
    @@current_sheet_name = ""
  end
  def open_excel_file(pfad)
    #dateiname = FILE_PATH + @excel_datei_name
    dateiname = pfad
    if File.exist?(dateiname) #(@excel_datei_name)
      #      @excel_appl = WIN32OLE.GetActiveObject('Excel.Application') ||
      #        WIN32OLE.new('Excel.Application')
      @excel_appl = WIN32OLE.new('Excel.Application') 
      @excel_appl['Visible'] = true      
      @excel_appl.Workbooks.Open(dateiname)
      @mappe = WIN32OLE.connect('Excel.Application').ActiveWorkbook #problem wenn alter excel task noch aktiv
    else
      raise "Error Message: " "Datei '#{dateiname}' nicht vorhanden " # eventuell neue abfrage
    end
  end
#  def find_excel_sheet
#    sheet_count = @excel_appl.Worksheets.count
#    for i in (1..sheet_count) do #@excel_appl.Worksheets.each) do
#      @sheet << @excel_appl.Worksheets(i).Name
#    end
#    return @sheet
#  end
#  def select_excel_sheet
#    if @sheet.include?(@excel_blatt_name)
#      @excel_appl.Worksheets(@excel_blatt_name).select
#      @@current_sheet_name = @excel_blatt_name
#      return @@current_sheet_name
#    else
#      raise "Error Message: " "Sheet nicht vorhanden" # eventuell neue abfrage
#    end
#  end
#  def close_excel_file()
#    @excel_appl.ActiveWorkbook.Close(0)
#  end
  def quit_excel
    @excel_appl.quit
    @excel_appl = nil
    GC.start #garbage control starten
  end
end

class ExcelLeser < ExcelController
  require 'win32ole'
  require 'stufen_tester'
  attr_accessor :column_name, :first_std_column
  def initialize(pfad, global_name, tabelle_name)
    @start_addresse = nil
    @zeilen_offset = 2
    i = 0
    @tabelle_name = tabelle_name
    super
    open_excel_file(pfad)
   end
  def zeile(zeilen_nummer)
    i=0
    loop do 
      @aktuelle_zeile = @excel_appl.WorkSheets(@tabelle_name).
        range("A1").offset(zeilen_nummer,i).value
      i+=1
      
      break if @aktuelle_zeile == nil
    end
  end
end

