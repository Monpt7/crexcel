require "./spec_helper"
require "../src/crexcel/*"

describe Crexcel do
  name = "test.xlsx"
  workbook = Crexcel::Workbook.new(name)
  worksheet1 = workbook.add_worksheet()

  it "creates a workbook with the correct filename" do
    workbook.name.should eq("test.xlsx")
  end

  it "creates a worksheet with an empty filename" do
    worksheet1.name.should eq("Sheet1")
  end

  it "creates a worksheet with the correct filename" do
    sheetname = "myname"
    worksheet2 = workbook.add_worksheet(sheetname).name
    worksheet2.should eq("myname")
  end

  it "creates a string" do
    Crexcel::SharedString.new("hi")
    Crexcel::SharedString.get_strings.size.should eq(1)
  end

  it "don't recreate the string if it exists" do
    Crexcel::SharedString.new("hi")
    Crexcel::SharedString.get_strings.size.should eq(1)
  end

  it "create a new string" do
    Crexcel::SharedString.new("how are you?")
    Crexcel::SharedString.get_strings.size.should eq(2)
  end

  it "write an int in cell" do
    worksheet1.write("A1", 123)
    worksheet1.get_datas.last["value"].should eq("123")
    worksheet1.get_datas.last["type"].should eq("")
    worksheet1.write("B1", 321)
  end

  it "write a float in cell" do
    worksheet1.write("A3", 2.2973)
    worksheet1.get_datas.last["value"].should eq("2.2973")
    worksheet1.get_datas.last["type"].should eq("")
  end

  it "write a string in cell" do
    worksheet1.write("A2", "hello")
    Crexcel::SharedString.get_strings.size.should eq(3)
    worksheet1.get_datas.last["type"].should eq("s")
  end

  workbook.close

end
