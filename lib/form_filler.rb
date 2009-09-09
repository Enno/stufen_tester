

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
    WIN32OLE.codepage = WIN32OLE::CP_UTF8 #zeichen als unicode verarbeiten
 
    p @xlapp.version
    case @xlapp.version
    when "12.0"
      @window_name = 'Microsoft Excel' #fuer office 07 anwendungen
      @access_to_macro = 1.0
    when "11.0", "10.0"
      @window_name = "Microsoft Excel - #{@file_name}" # für Office XP/2002
      @access_to_macro = 0.5
      p "schnelles Excel"
    end
  end

  def open_template
    @template_controller.sende_tasten(@window_name, "%{F8}#{@proc_name}%{a}", :wartezeit => 0.2, :fenster_fehlt=>"Komischerweise fehlt das Excel-Fenster")
    sleep(@access_to_macro||0.7)
  end

  def tab_set(numbers = 1)
    return unless numbers
    shift_code = numbers < 0 ? "+" : ""
    send_tabs = "#{shift_code}{TAB}" * numbers.abs
    send_keys("#{send_tabs}", :wartezeit => 0.05)
  end

  def confirm_input
    send_keys("{ENTER}", :wartezeit => 0.2)
  end

  def send_keys(character, options = {})
    @template_controller.sende_tasten(@fenstername, "#{character}", options)
  end

  SUM               = proc {|a, b| a + b}
  GREATER_THEN_ZERO = proc {|a| a > 0}

  @@records = [
    [
      :name,
      :bruttogehalt,
      :freibetrag,
      {:k_vers_art => {
          :nature           => :radio_group,
          :select_list      => ["g", "p"],
          :default_value    => "g",
          :deactivates      => [:kinderlos]
        }},
      {:steuerklasse=> {
          :nature             => :direkt,
          :select_list        => ["I", "II", "III", "IV", "V", "VI"],
          :default_value      => ["V", "VI"],
          :activates          => [:kinder_fb],
          :skip_adjustment    => 0
        }},
      :kinder_fb,
      {:kirchensteuer       => true},
      :bland_wohnsitz,
      :bland_arbeit,
      :berufsgruppe,
      {:durchfuehrungsweg   => {
          :nature             => :direkt,
          :select_list        => ["Direktversicherung", "Pensionskasse", "Unterstützungskasse"],
          :default_value      => "Unterstützungskasse",
          :activates          => [:pausch_steuer40b],
          :skip_adjustment    => 0
        }},
      {:pausch_steuer40b    => false},
      {:minijob_ok          => false},
      {:kinderlos           => false},
    ],[
      {:nvz => {
          :nature           => :checkbox,
          :default_value    => true,
          :deactivates      => [:verzicht_betrag, :verzicht_als_netto],
          :function         => GREATER_THEN_ZERO,
          :params           => [:verzicht_betrag]
        }},
      :verzicht_betrag,
      {:verzicht_als_netto  => {
          :nature           => :radio_group,
          :select_list      => [true, false], #["netto", "brutto"]
          :default_value    => true
        }},
    ],[
      :vl_arbeitgeber,
      {:vl_gesamt => {
          :nature           => :direkt,
          :function         => SUM,
          :params           => [:vl_arbeitgeber, :vl_arbeitnehmer],
        }},
      {:vl_als_beitrag      => true}
    ],[
      {:ag_zuschuss_ok => {
          :nature           => :checkbox,
          :default_value    => false,
          :skip_adjustment  => -3,
          :activates        => [:ag_zuschuss, :ag_zuschuss_als_absolut],
          :function         => GREATER_THEN_ZERO,
          :params           => [:ag_zuschuss]

        }},
      {:ag_zuschuss_als_absolut => {
          :nature           => :radio_group,
          :select_list      => [true, false], # ["€", "%"],
          :default_value    => true,          # "€",
          :skip_adjustment  => 0
        }},
      :ag_zuschuss
    ]
  ]

  def identify_nature(dataset, symbol_or_hash)
    case symbol_or_hash
    when Symbol
      @nature              = :direkt
      @actual_box_name     = symbol_or_hash
    when Hash
      @right_side = symbol_or_hash.values.first
      is_complex  = @right_side.is_a?(Hash)
      @nature     = case @right_side
      when Array            then :radio_group
      when true, false      then :checkbox
      when Hash             then @right_side[:nature]
      end
      @actual_box_name    = symbol_or_hash.keys.first
    end    
    select_processing_data(dataset, is_complex)
  end

  def select_processing_data(dataset, is_complex)
    @continue_processing_data = if is_complex and @right_side[:function]
      param_values  = @right_side[:params].map {|symbol| dataset[symbol] }
      function      = @right_side[:function]
      function[*param_values]
    else
      dataset[@actual_box_name]
    end
    check_variables_environment(is_complex)
  end

  def check_variables_environment(is_complex)
    if is_complex then
      skip_adjustment     = @right_side[:skip_adjustment]
      select_list         = @right_side[:select_list]
      default_value       = @right_side[:default_value]
      if default_value.is_a?(Array)
        default_value.each do |value|
          non_busy_boxes_new = (@continue_processing_data != value ?
              @right_side[:deactivates] :
              @right_side[:activates])
          @non_busy_boxes     += non_busy_boxes_new if non_busy_boxes_new
        end
      else
        non_busy_boxes_new = (@continue_processing_data != default_value ?
            @right_side[:deactivates] :
            @right_side[:activates])
        @non_busy_boxes     += non_busy_boxes_new if non_busy_boxes_new
      end
    else
      skip_adjustment     = 0
      case @nature
      when :checkbox
        select_list       = nil
        default_value     = @right_side
      when :radio_group
        select_list       = @right_side
        default_value     = nil
        nil
      end
    end
    return @processing_data_attributes = {
      "actual_box_name" => @actual_box_name,
      "nature"          => @nature,
      "select_list"     => select_list,
      "default_value"   => default_value,
      "skip_adjustment" => skip_adjustment,
      "non_busy_boxes"  => @non_busy_boxes}
  end

  def enter_value(dataset, box_info)

    identify_nature(dataset, box_info)

    return if @processing_data_attributes["non_busy_boxes"].include? @processing_data_attributes["actual_box_name"]

    case @processing_data_attributes["nature"]
    when :direkt
      send_keys(@continue_processing_data.is_a?(Float) ?
          (change_decimal_seperation(@continue_processing_data)) : (@continue_processing_data))
      tab_set
    when :checkbox
      send_keys(' ') if @processing_data_attributes["default_value"] ^ @continue_processing_data # exclusive or
      tab_set
    when :radio_group
      change = (@continue_processing_data != @processing_data_attributes["default_value"])
      @processing_data_attributes["select_list"].each do |feasible_value|
        if change and feasible_value == @continue_processing_data then
          send_keys(' ')
          break if @processing_data_attributes["skip_adjustment"]
        end
        tab_set
      end
    end
    tab_set(@processing_data_attributes["skip_adjustment"])
  end

  def change_decimal_seperation(continue_processing_data)
    return continue_processing_data.to_s.gsub(/[.]/, ',')
  end

  def populate_template(dataset)
    open_template
    tab_index = 1
    @@records.each do |boxes_in_actual_tab|
      @processing_data_attributes = {}
      @non_busy_boxes             = []
      boxes_in_actual_tab.each do |box_info|
        enter_value(dataset, box_info)
      end
      break if tab_index == 4
      send_keys('^{PGDN}')
      tab_index += 1
    end
    start_calculation
  end

  def start_calculation #besser waere es, wenn der button "ergebnis" direkt angesprochen werden kann
    tab_set(15)  #sichergehen, dass ergebnis-button erreicht wird
    tab_set(-2)
    sleep 1
    confirm_input
    sleep(0.2)
    confirm_input
  end

  def vb_send(vb_procedure_name, *args)
    @xlapp.Run "#{@file_name}!#{vb_procedure_name}", *args
  end

  def close_template #ueber button "schliessen" siehe kommentar "start_calculation"
    @xlapp.ActiveWorkbook.Close
  end

  def quit_excel
    @excel_controller.quit_excel
  end

end