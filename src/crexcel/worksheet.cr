require "../helpers/position_helper.cr"

module Crexcel
  # TODO Write doc for Worksheet class
  class Worksheet
    alias Datum = NamedTuple(pos: String, value: String, type: String)

    SUPPORTED_TYPES = {"string" => "s", "int" => "n"}
    @datas : Array(Datum)

    # Get name of the worksheet
    getter name : String

    # :nodoc:
    def initialize(name : String)
      @datas = Array(Datum).new
      @name = name
    end

    # Write data in the specified worksheet at specific location
    #
    # ```
    # first_worksheet = workbook.add_worksheet
    # first_worksheet.write("A1", "Hello world!")
    # first_worksheet.write("A2", "1337") # Will write 1337 as a string
    # first_worksheet.write("B1", 1337)   # Will write 1337 as a number
    # ```
    def write(position : String, str : Int32 | Int64 | Float64 | String)
      if str.is_a? String
        value = Crexcel::SharedString.new(str).index.to_s
        type = "s"
      else
        value = str.to_s
        type = ""
      end
      @datas << {pos: position, value: value, type: type}
    end

    # Write data in the specified worksheet at specific location
    # You can precise the type you want
    #
    # ```
    # first_worksheet = workbook.add_worksheet
    # first_worksheet.write("A1", "1337", "int") # Will write 1337 as a number
    # ```
    # **Available types are for now :** string, int
    def write(position : String, str : Int32 | Int64 | Float64 | String, type_req : String)
      raise "Unsupported type for write method" unless SUPPORTED_TYPES.has_key? type_req
      type = SUPPORTED_TYPES[type_req]
      value = type == "s" && str.is_a?(String) ? Crexcel::SharedString.new(str).index.to_s : str.to_s
      @datas << {pos: position, value: value, type: type}
    end

    # Same as other write methods, but location is precised differently
    # ```
    # first_worksheet = workbook.add_worksheet
    # first_worksheet.write(0, 0, "hello") # Will write "hello" in A1 cell
    # ```
    def write(pos_x : Int32 | Int64, pos_y : Int32 | Int64, str : Int32 | Int64 | Float64 | String, type_req = "")
      if type_req == ""
        write(int_pos_to_char(pos_x, pos_y), str)
      else
        write(int_pos_to_char(pos_x, pos_y), str, type_req)
      end
    end

    # Will write a row with the provided array of data
    # ```
    # first_worksheet = workbook.add_worksheet
    # first_worksheet.write_row(0, ["hello", "world"]) # Will write "hello" in A1 cell and "world" in A2
    # ```
    def write_row(pos_y : Int32 | Int64, data : Array(Int32 | Int64 | Float64 | String))
      data.each_index do |i|
        write(i, pos_y, data[i])
      end
    end

    # Will write a row with the provided array of data
    # ```
    # first_worksheet = workbook.add_worksheet
    # first_worksheet.write_row("A", ["hello", "world"]) # Will write "hello" in A1 cell and "world" in A2
    # ```
    def write_row(row : String, data : Array(Int32 | Int64 | Float64 | String))
      data.each_index do |i|
        position = "#{row}#{i + 1}"
        write(position, data[i])
      end
    end

    # Fetch the data at the given string postition
    # ```
    # first_worksheet = workbook.add_worksheet
    # first_worksheet.write(0, 0, "hello") # Will write "hello" in A1 cell
    # first_worksheet.get("A1")            # => {pos: "A1", value: "hello", type: ""}
    # ```
    def get(position : String) : Datum
      get_datas.find { |i| i[:pos] == position } || {pos: position, value: "", type: ""}
    end

    # Fetch the data at the given x, y postition
    # ```
    # first_worksheet = workbook.add_worksheet
    # first_worksheet.write(0, 0, "hello") # Will write "hello" in A1 cell
    # first_worksheet.get(0, 0)            # => {pos: "A1", value: "hello", type: ""}
    # ```
    def get(pos_x : Int32 | Int64, pos_y : Int32 | Int64)
      position = int_pos_to_char(pos_x, pos_y)
      get_datas.find { |i| i[:pos] == position } || {pos: position, value: "", type: ""}
    end

    # :nodoc:
    def get_datas
      @datas
    end

    # :nodoc:
    def get_tidy_datas
      tidy_datas = Hash(Int32, Array(Datum)).new
      sorted_datas = @datas.sort { |a, b|
        tmpa = /[A-Z]*([\d]*)/.match(a["pos"])
        tmpb = /[A-Z]*([\d]*)/.match(b["pos"])
        if tmpa.not_nil![1].to_i > tmpb.not_nil![1].to_i
          1
        elsif tmpa.not_nil![1].to_i == tmpb.not_nil![1].to_i
          if tmpa.not_nil![0] > tmpb.not_nil![0]
            1
          else
            -1
          end
        else
          -1
        end
      }
      sorted_datas.each do |data|
        row = /[A-Z]*([\d]*)/.match(data["pos"])
        row = row.not_nil![1].to_i
        tidy_datas[row] = Array(Datum).new unless tidy_datas.has_key?(row)
        tidy_datas[row] << data
      end
      tidy_datas
    end
  end
end
