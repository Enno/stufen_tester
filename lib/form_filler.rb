

require 'win32ole'

require 'grundlage'
require 'excel_controller'
require 'tasten_sender'

#constant definition
#FILE_PATH = File.dirname(File.dirname(__FILE__)).gsub('\\', '/') +  '/daten/'

class FormFiller 
 
  def initialize(path, file_name, start_proc_name)
    @file_name = file_name
    @proc_name = start_proc_name
    @excel_controller = ExcelController.new(path + file_name)
    @excel_controller.open_excel_file(path + file_name)
    @xlapp = @excel_controller.excel_appl
    @template_controller = TastenSender.new(:wartezeit => 0.2)
 
    case @xlapp.version
    when "12.0"
      @window_name = 'Microsoft Excel' #fuer office 07 anwendungen
      @access_to_macro = 1.0
    when "11.0"
      @window_name = "Microsoft Excel - #{@file_name}" # für Office XP/2002
      @access_to_macro = 0.5
    end
  end

  def open_template
    @template_controller.sende_tasten(@window_name, "%{F8}#{@proc_name}%{a}", :wartezeit => 0.2, :fenster_fehlt=>"Komischerweise fehlt das Excel-Fenster")
    sleep(@access_to_macro)
  end

  def tab_set(numbers = 1)
    return unless numbers
    shift_code = numbers < 0 ? "+" : ""
    send_tabs = "#{shift_code}{TAB}" * numbers.abs
    send_keys("#{send_tabs}", :wartezeit => 0.01)
  end

  def confirm_input
    send_keys("{ENTER}", :wartezeit => 0.2)
  end

  def send_keys(character, options = {})
    @template_controller.sende_tasten(@fenstername, "#{character}", options)
  end

  SUM = proc {|a, b| a + b}
  @@records = [
    [
      :name,
      :bruttogehalt,
      :freibetrag,
      {:k_vers_nature => ["g","p"]},
      :steuerklasse,
      :kinder_fb,
      {:kirchensteuer => true},
      :bland_wohnsitz,
      :bland_arbeit,
      :berufsgruppe,
      :durchfuehrungsweg,
      {:pausch_steuer40b => false},
      {:minijob_ok => false},
      {:kinderlos => false},
    ],[
      {:umwandlgvon_keine_ahnung_welches_box => false},
      :verzicht_betrag,
      {:verzicht_als_netto => ["netto", "brutto"]}
    ],[
      :vl_arbeitgeber,
      {:vl_gesamt => {
          :nature => :direkt,
          :function => SUM,
          :params => [:vl_arbeitgeber, :vl_arbeitnehmer],
        }},
      {:vl_als_beitrag => true}
    ],[
      {:ag_zuschuss_ok => {
          :nature => :checkbox,
          :default_value => false,
          :skip_adjustment => -3,
          :activates => [:ag_zuschuss, :ag_zuschuss_als_absolut]
        }},
      {:ag_zuschuss_als_absolut => {
          :nature => :radio_group,
          :select_list => ["€", "%"],
          :default_value => "€",
          :skip_adjustment => 0
        }},
      :ag_zuschuss
    ]
  ]

  def identify_value(datset, symbol_or_hash)
    case symbol_or_hash
    when Symbol
      nature = :direkt
      sym = symbol_or_hash
    when Hash
      right_side = symbol_or_hash.values.first
      is_complex = right_side.is_a?(Hash)
      nature = case right_side
      when Array            then :radio_group
      when true, false      then :checkbox
      when Hash             then right_side[:nature]
      end
      sym = symbol_or_hash.keys.first
    end

    return  if @non_busy_boxes.include? sym

    continue_processing_data = if is_complex and right_side[:function]
      param_values = right_side[:params].map {|symbol| datset[symbol] }
      function = right_side[:function]
      function[*param_values]
    else
      datset[sym]
    end
    preparing_value(nature, right_side, continue_processing_data)
  end

  # Vor-Verarbeitung
  def preparing_value(nature, right_side, continue_processing_data)
    @default_value = if right_side.is_a?(Hash) then
      @skip_adjustment = right_side[:skip_adjustment]
      @select_list = right_side[:select_list]

      non_busy_boxes_new = (continue_processing_data ? [] : right_side[:activates])
      @non_busy_boxes += non_busy_boxes_new if non_busy_boxes_new

      right_side[:default_value]
    else
      @skip_adjustment = 0
      case nature
      when :checkbox
        right_side
      when :radio_group
        @select_list = right_side
        nil
      end
    end
    enter_value(continue_processing_data, nature)
  end

  def enter_value(continue_processing_data, nature)
    case nature
    when :direkt
      send_keys(continue_processing_data.is_a?(Float) ?
          (change_decimal_seperation(continue_processing_data)) : (continue_processing_data))
      tab_set
    when :checkbox
      send_keys(' ') if @default_value ^ continue_processing_data # exclusive or
      tab_set
    when :radio_group
      change = (continue_processing_data != @default_value)
      @select_list.each do |feasible_value|
        if change and feasible_value == continue_processing_data then
          send_keys(' ')
          break if @skip_adjustment
        end
        tab_set
      end
    end
    tab_set( @skip_adjustment )
  end

  def change_decimal_seperation(continue_processing_data)
    return continue_processing_data.to_s.gsub(/[.]/, ',')
  end

  def populate_template(datset)
    open_template
    tab_index = 1
    @@records.each do |boxes_in_actual_tab|
      @non_busy_boxes = []
      boxes_in_actual_tab.each do |box_info|
        identify_value(datset, box_info)
      end
      break if tab_index == 4
      send_keys('^{PGDN}')
      tab_index += 1
    end
    start_calculation
  end

  def start_calculation #besser waere es, wenn der button "ergebnis" direkt angesprochen werden kann
    sleep 2
    confirm_input
    sleep(0.1)
    confirm_input
  end

  def vb_send(vb_request)
    @xlapp.Run "#{@file_name}!#{vb_request}"
  end

  def close_template #ueber button "schliessen" siehe kommentar "start_calculation"
    @xlapp.ActiveWorkbook.Close
  end

  def quit_excel
    @excel_controller.quit_excel
  end

end

#@todo: berechnetes box (groesser als 0 etc)
#@todo: code refactoring bsp: tastensenden vor if
#@todo: strg pagedown fuer registerknatureen wechseln
#@todo: excel_leser anbindung