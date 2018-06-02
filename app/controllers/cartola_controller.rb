class CartolaController < ApplicationController
  require 'roo'
  require 'spreadsheet'
  require 'roo-xls'


  def index

  end

  # Summary: Lectura del archivo
  #
  # Params: Archivo que se va a procesar
  # Return: status
  def import
    begin
      movements = Roo::Excelx.new(params[:file].path, file_warning: :ignore)
      #Roo::Spreadsheet.open(params[:file], extension: :xlsx)
      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")
        if File.file?(export_file_path)
          validate_movement(movements)
        else
          create_movement(movements)
        end
      return render json: {status: true}
    rescue Exception => e  
      puts e
      return false
    end
  end 

  # Summary: Validacion de cartola
  #
  # Params:
  # Return:  
  def validate_movement movements
    begin  
      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")
      movements_old = Roo::Spreadsheet.open(export_file_path, extension: :xls)
      #movements_old = Spreadsheet.open(export_file_path)

      binding.pry
      return render json: {status: true, header: sheet1}
    rescue Exception => e 
      puts e
      return render json: {status: false}
    end
  end 

  # Summary: Creacion de cartola correcta
  #
  # Params:
  # Return:  
  def create_movement movements
    begin  
      header = movements.row(1)
      book = Spreadsheet::Workbook.new 
      sheet1 = book.create_worksheet :name => 'cartola_correcta'
      sheet1.row(0).replace header
      count = 0
      (2..movements.last_row).each do |i|
        count = count + 1
        sheet1.row(count).replace movements.row(i)
      end

      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")
      book.write export_file_path
      send_file export_file_path, :content_type => "application/vnd.ms-excel", :disposition => 'inline'

      return render json: {status: true, header: sheet1}
    rescue Exception => e 
      puts e
      return render json: {status: false}
    end
  end
end