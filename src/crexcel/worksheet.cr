require "../helpers/position_helper.cr"
module Crexcel
  # TODO Write doc for Worksheet class
  class Worksheet
    @datas : Array(NamedTuple(pos: String, value: String, type: String))
    getter name : String

    def initialize(name : String)
      @datas = Array(NamedTuple(pos: String, value: String, type: String)).new
      @name = name
    end

    def write(position : String, str : Int32 | Int64 | Float64 | String )
      if str.is_a? String
        value = Crexcel::SharedString.new(str).index.to_s
        type = "s"
      else
        value = str.to_s
        type = ""
      end
      @datas << {pos: position, value: value, type: type}
    end

    def write(pos_x : Int32 | Int64, pos_y : Int32 | Int64, str : Int32 | Int64 | Float64 | String )
      write(int_pos_to_char(pos_x, pos_y), str)
    end

    def get_datas
      @datas
    end

    def get_tidy_datas
      tidy_datas = Hash(Int32, Array(NamedTuple(pos: String, value: String, type: String))).new
      sorted_datas = @datas.sort {|a,b|
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
        tidy_datas[row] = Array(NamedTuple(pos: String, value: String, type: String)).new unless tidy_datas.has_key?(row)
        tidy_datas[row] << data
      end
      tidy_datas
    end

  end
end
