require 'test_helper'

class NavigationTest < ActiveSupport::IntegrationCase
  test "Render index.xls return xls file" do
    visit root_path(:format => :xls)
    assert_equal 'binary', page.response_headers['Content-Transfer-Encoding']
    assert_equal 'application/vnd.ms-excel', page.response_headers['Content-Type']
  end
  
  test "Render index.xlsx return xlsx file" do
    visit root_path(:format => :xlsx)
    assert_equal 'binary', page.response_headers['Content-Transfer-Encoding']
    assert_equal 'application/vnd.openxmlformats', page.response_headers['Content-Type']
  end
  
  test "Render export.xls return xls file" do
    visit export_items_path(:format => :xls)
    assert_equal 'binary', page.response_headers['Content-Transfer-Encoding']
    assert_equal 'application/vnd.ms-excel', page.response_headers['Content-Type']
  end
  
   
end
