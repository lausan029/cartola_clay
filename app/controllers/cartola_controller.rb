class CartolaController < ApplicationController
  require 'roo'
  require 'spreadsheet'
  require 'roo-xls'

  include CartolaHelper

  def index
  end

  # Summary: Lectura del archivo
  #
  # Params: Archivo que se va a procesar
  # Return: status
  def import
    begin
      movements        = Roo::Excelx.new(params[:file].path, file_warning: :ignore)
      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")

        if File.file?(export_file_path)
          movements_old = Roo::Spreadsheet.open(export_file_path, extension: :xls)
          if movements.last_row < movements_old.last_row
            especial_movement(movements,movements_old)
          else
            validate_movement(movements,movements_old)
          end
        else
          create_movement(movements)
        end
      return render json: {status: true}
    rescue Exception => e
      puts e
      return false
    end
  end 
end