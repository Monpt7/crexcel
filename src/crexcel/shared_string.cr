module Crexcel
  # :nodoc:
  class SharedString
    @@strings = Array(String).new
    @@count = 0
    getter index : Int32

    def initialize(string : String)
      unless @@strings.includes?(string)
        @index = @@strings.size.to_i
        @@strings << string
        @@count += 1
      else
        ind = @@strings.index(string)
        @index = ind.nil? ? @@strings.size.to_i : ind
      end
    end

    def self.get_strings
      @@strings
    end

    def self.generate_xml(file : String)
      string = XML.build(encoding: "UTF-8") do |xml|
        xml.element("sst", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", count: @@count.to_s, uniqueCount: @@strings.size.to_s) do
          @@strings.each do |str|
            xml.element("si") do
              xml.element("t") { xml.text str }
            end
          end
        end
      end
      File.write(file, string)
    end
  end
end
