class CartolaController < ApplicationController

  def index
  end
  require 'roo'
  def import
    binding.pry
    cartola =  Roo::Excelx.new(params[:file].path, file_warning: :ignore)
    puts cartola
    page_cartola = cartola.sheets
    page_cartola.each do |c|
    	cartola.sheet(cartola.sheets.firts).each_row_streaming do |row|
    		row_cells = row.map { |cell| puts cell.value }
    	end
    end
    
    redirect_to root_url, notice: 'Products imported.'
  end
end

=begin
	
	column(column_number, sheet = nil)

rescue Exception => e
	
end
workbook = Roo::Spreadsheet.open './sample_excel_files/xlsx_500_rows.xlsx'
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet}"
  num_rows = 0
  workbook.sheet(worksheet).each_row_streaming do |row|
    row_cells = row.map { |cell| puts cell.value }
    num_rows += 1
  end
  puts "Read #{num_rows} rows" 
end

worksheets.each do |worksheet|
puts "Reading: #{worksheet}"
end

worksheets.each do |worksheet|
workbook.sheet(worksheet).each_row_streaming do |row|
    puts row
  end
end
=end