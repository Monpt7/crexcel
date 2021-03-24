require "../ext/dir"
require "compress/zip"
require "./worksheet"
require "file_utils"

module Crexcel
  # The Workbook Class is the main class of Crexcel.
  # It create the file which will contains the worksheets.
  # NOTE: Don't forget to call the close method in order to generate your xlsx file!
  class Workbook
    alias Dirs = NamedTuple(
      root: String,
      xl: String,
      rels: String,
      doc_props: String,
      xl_rels: String,
      xl_theme: String,
      worksheets: String,
    )

    @sheets : Array(Worksheet)
    @directory : String
    @tmpdir : String = Dir.tempdir

    getter name : String
    getter dirs : Dirs

    # The new method let you create a workbook, you only need one per file you want to generate.
    #
    # You must give a name to it :
    #
    # ```
    # my_workbook = Workbook.new("thing.xlsx")
    # ```
    # NOTE: You are not forced to give the extension, if you type `Workbook.new("thing")`
    #  it will still create a 'thing.xlsx' file.
    def initialize(name : String)
      @sheets = Array(Worksheet).new
      @name = name
      raise "Error: Temporary directory doesn't exist" unless Dir.exists?(@tmpdir)
      @directory = Dir.mktmpdir

      @dirs = {
        root:       @directory,
        xl:         File.join(@directory, "xl"),
        rels:       File.join(@directory, "_rels"),
        doc_props:  File.join(@directory, "docProps"),
        xl_rels:    File.join(@directory, "xl", "_rels"),
        xl_theme:   File.join(@directory, "xl", "theme"),
        worksheets: File.join(@directory, "xl", "worksheets"),
      }
    end

    # Call this to add a worksheet to the workbook without giving it a name.
    #
    # It returns a Worksheet object
    # ```
    # workbook = Workbook.new("thing.xlsx")
    # first_worksheet = workbook.add_worksheet
    # ```
    def add_worksheet
      nbr_sheet = @sheets.size + 1
      name = "Sheet" + nbr_sheet.to_s
      worksheet = Worksheet.new(name)
      @sheets << worksheet
      worksheet
    end

    # Call this to add a worksheet to the workbook with a name.
    #
    # It returns a Worksheet object
    # ```
    # workbook = Workbook.new("thing.xlsx")
    # first_worksheet = workbook.add_worksheet("test")
    # ```
    def add_worksheet(name : String)
      worksheet = Worksheet.new(name)
      @sheets << worksheet
      worksheet
    end

    # This function close the workbook and generate a xlsx file.
    # You **MUST** call this function at the end of you script for each Workbook created.
    # If you don't, this **WILL NOT** generate the xlsx file.
    # OPTIMIZE: In the future, not closing a Workbook will raise an error
    #  instead of silently not generate the file.
    def close
      dirs.each_value do |val|
        Dir.mkdir(val) unless Dir.exists?(val)
      end

      ArchiveBuilder.new(dirs, @sheets.size)
      SharedString.generate_xml(File.join(dirs[:xl], "sharedStrings.xml"))
      write_worksheets_xml
      write_workbook_xml
      write_rels_xml
      generate_xlsx
      FileUtils.rm_r(dirs[:root])
    end

    private def generate_xlsx
      name = @name
      name = name + ".xlsx" if name.split('.')[-1] != "xlsx"
      File.open(name, "w") do |file|
        Compress::Zip::Writer.open(file) do |zip|
          zip.add("_rels/.rels", File.open(File.join(dirs[:rels], ".rels")))
          zip.add("docProps/app.xml", File.open(File.join(dirs[:doc_props], "app.xml")))
          zip.add("docProps/core.xml", File.open(File.join(dirs[:doc_props], "core.xml")))
          zip.add("xl/_rels/workbook.xml.rels", File.open(File.join(dirs[:rels], "workbook.xml.rels")))
          zip.add("xl/theme/theme1.xml", File.open(File.join(dirs[:xl], "theme", "theme1.xml")))
          @sheets.each_with_index do |sheet, i|
            i += 1
            zip.add("xl/worksheets/sheet#{i}.xml", File.open(File.join(dirs[:worksheets], "sheet#{i}.xml")))
          end
          zip.add("xl/sharedStrings.xml", File.open(File.join(dirs[:xl], "sharedStrings.xml")))
          zip.add("xl/styles.xml", File.open(File.join(dirs[:xl], "styles.xml")))
          zip.add("xl/workbook.xml", File.open(File.join(dirs[:xl], "workbook.xml")))
          zip.add("[Content_Types].xml", File.open(File.join(dirs[:root], "[Content_Types].xml")))
        end
      end
    end

    private def write_workbook_xml
      string = XML.build(encoding: "UTF-8") do |xml|
        xml.element("workbook", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships") do
          xml.element("fileVersion", appName: "xl", lastEdited: "1", lowestEdited: "1", rupBuild: "9303")
          xml.element("workbookPr", defaultThemeVersion: "124226")
          xml.element("bookViews") do
            xml.element("workbookView", xWindow: "240", yWindow: "105", windowWidth: "9660", windowHeight: "16095")
          end
          xml.element("sheets") do
            @sheets.each_with_index do |sheet, i|
              i += 1
              xml.element("sheet", name: sheet.name, sheetId: i, "r:id": "rId#{i}")
            end
          end
          xml.element("calcPr", calcId: "124519", fullCalcOnLoad: "1")
        end
      end
      File.write(File.join(dirs[:xl], "workbook.xml"), string)
    end

    private def write_worksheets_xml
      @sheets.each_with_index do |sheet, i|
        datas = sheet.get_tidy_datas
        i += 1
        string = XML.build(encoding: "UTF-8") do |xml|
          xml.element("worksheet", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships") do
            xml.element("dimension", ref: "A1:A4")
            xml.element("sheetViews") do
              xml.element("sheetView", tabSelected: "1", workbookViewId: "0")
            end
            xml.element("sheetFormatPr", defaultRowHeight: "15")
            xml.element("sheetData") do
              datas.each do |key, value|
                xml.element("row", r: key.to_s, spans: "1:1") do
                  value.each do |data|
                    if data["type"] == ""
                      xml.element("c", r: data["pos"]) do
                        xml.element("v") { xml.text data["value"] }
                      end
                    else
                      xml.element("c", r: data["pos"], t: data["type"]) do
                        xml.element("v") { xml.text data["value"] }
                      end
                    end
                  end
                end
              end
            end
            xml.element("pageMargins", left: "0.7", right: "0.7", top: "0.75", bottom: "0.75", header: "0.3", footer: "0.3")
          end
        end
        File.write(File.join(dirs[:worksheets], "sheet#{i}.xml"), string)
      end
    end

    private def write_rels_xml
      worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      r_id = 1
      string = XML.build(encoding: "UTF-8") do |xml|
        xml.element("Relationships", xmlns: "http://schemas.openxmlformats.org/package/2006/relationships") do
          @sheets.each do |sheet|
            xml.element("Relationship", "Id": "rId#{r_id}", "Type": worksheet_type, "Target": "worksheets/sheet#{r_id}.xml")
            r_id += 1
          end
          xml.element("Relationship", "Id": "rId#{r_id}", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "Target": "theme/theme1.xml")
          xml.element("Relationship", "Id": "rId#{r_id + 1}", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "Target": "styles.xml")
          xml.element("Relationship", "Id": "rId#{r_id + 2}", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "Target": "sharedStrings.xml")
        end
      end
      File.write(File.join(dirs[:rels], "workbook.xml.rels"), string)
    end
  end
end
