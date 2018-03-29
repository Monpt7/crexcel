module Crexcel
  # TODO Write doc for Worksheet class
  class Worksheet
    @datas : Array(NamedTuple(pos: String, value: String, type: String))
    getter name : String

    def initialize(name : String)
      @datas = Array(NamedTuple(pos: String, value: String, type: String)).new
      @name = name
    end

    def write(position : String, str : Int32 | Int64 | String )
      if str.is_a? String
        value = Crexcel::SharedString.new(str).index.to_s
        type = "s"
      else
        value = str.to_s
        type = ""
      end
      @datas << {pos: position, value: value, type: type}
    end

    def get_datas
      @datas
    end

  end
end
