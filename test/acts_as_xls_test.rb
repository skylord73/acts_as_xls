require 'test_helper'

class ActsAsXlsTest < ActiveSupport::TestCase
  
  setup do
    @array = [[1,2,3],["a","b","c"], [Time.now.yesterday, Time.now, Time.now.tomorrow], [1.1, 2.2, 3.3]]
    @active = Item
  end
  
  
  test "Respond to new_from_xls" do
    assert Item.respond_to?(:new_from_xls)
  end

  test "Respond to to_xls" do
    assert Item.new.respond_to?(:to_xls)
  end
  
  test "xls mime type" do
    assert_equal :xls, Mime::XLS.to_sym
    assert_equal "application/vnd.ms-excel", Mime::XLS.to_s
  end
  
  test "xlsx mime type" do
    assert_equal :xlsx, Mime::XLSX.to_sym
    assert_equal "application/vnd.openxmlformats", Mime::XLSX.to_s
  end
  
  
  
  test "xls from array" do
    assert_nothing_raised do
      @spreadsheet = Spreadsheet.open(StringIO.new(@array.to_xls))
    end
    assert @spreadsheet.worksheet(0).last_row_index == (@array.length - 1), "Last index = #{@spreadsheet.worksheet(0).last_row_index}, length=#{@array.length}"
    
  end
  
  test "xls from array of ActiveRecord" do
    assert_nothing_raised do
      @spreadsheet = Spreadsheet.open(StringIO.new(@active.all.to_xls))
    end
    assert @spreadsheet.worksheet(0).last_row_index == (@active.all.length), "Last index = #{@spreadsheet.worksheet(0).last_row_index}, length=#{@active.all.length}"
    
  end
  
  test "xls from Relation" do
    assert_nothing_raised do
      @spreadsheet = Spreadsheet.open(StringIO.new(@active.where("speed > 0").to_xls))
    end
    assert @spreadsheet.worksheet(0).last_row_index == (@active.where("speed > 0").length), "Last index = #{@spreadsheet.worksheet(0).last_row_index}, length=#{@active.where("speed > 0").length}"
  end
  
  test "xls from ActiveRecord" do
    assert_nothing_raised do
      @spreadsheet = Spreadsheet.open(StringIO.new(@active.first.to_xls))
    end
    assert @spreadsheet.worksheet(0).last_row_index == 1, "Last index = #{@spreadsheet.worksheet(0).last_row_index}, length=1"
  end
  
  test "xls with custom columns" do
    assert_nothing_raised do
      @spreadsheet = Spreadsheet.open(StringIO.new(@active.all.to_xls(:only => ["name", "user.name", "fake_column"])))
    end
    assert @spreadsheet.worksheet(0).last_row_index == @active.all.length, "Last index = #{@spreadsheet.worksheet(0).last_row_index}, length =#{@active.all.length}"
  end

end
