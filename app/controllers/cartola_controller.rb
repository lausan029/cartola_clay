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
      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")
        if File.file?(export_file_path)
          movements_old = Roo::Spreadsheet.open(export_file_path, extension: :xls)
            if movements.last_row < movements_old.last_row
              especial_movement(movements,movements_old)
            else
              validate_movement(movements)
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

  # Summary: Validación de cartola cuando es cambio de mes
  #
  # Params: Archivo precesado anteriormente
  # Return: status
  def especial_movement movements,movements_old
    begin
      book = Spreadsheet::Workbook.new
      sheet1 = book.create_worksheet :name => 'cartola_correcta'
      sheet1.row(0).replace movements_old.row(1)
      count = 0
        #Llena sheet con archivo viejo
        (2..movements_old.last_row).each do |i|
          count = count + 1
          sheet1.row(count).replace movements_old.row(i)
        end
        #remplazo repetido viejos por actualizados
        (2..movements.last_row).each do |rw|
          (2..movements_old.last_row).each do |r|
            if (movements.row(rw).last == movements_old.row(r).last) and (movements.row(rw).first != movements_old.row(r).first)
              sheet1.row(r).replace movements.row(rw)
            end
          end
        end
        #Inserta resto de archivo nuevo
        (2..movements.last_row).each do |rw|
          unless movements.row(rw).presence_in sheet1
            count = sheet1.row_count + 1
            sheet1.row(count).replace movements.row(rw)
          else
            puts "existe"
          end
        end
        #Quito los rows nil
        (1..sheet1.row_count).each do |rm|
          unless sheet1.row(rm).present?
             sheet1.delete_row(rm)
          end
        end

      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")
      book.write export_file_path
      send_file export_file_path, :content_type => "application/vnd.ms-excel", :disposition => 'inline'

      return render json: {status: true}
    rescue Exception => e
      puts e
      return render json: {status: false}
    end
  end

  # Summary: Validación de cartola
  #
  # Params: Archivo precesado anteriormente
  # Return: status
  def validate_movement movements
    begin  
      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")
      movements_old = Roo::Spreadsheet.open(export_file_path, extension: :xls)

      book = Spreadsheet::Workbook.new
      sheet1 = book.create_worksheet :name => 'cartola_correcta'
      count = 0

      (1..movements.last_row).each do |i|
        (1..movements_old.last_row).each do |r|
          (1..movements.last_row).each do |rw|
            if movements_old.row(r) == movements.row(rw)
              sheet1.row(count).replace movements.row(i)
            else
              if (movements_old.row(r).last.to_i == movements.row(rw).last) and (movements_old.row(r).first != movements.row(r).first)
                sheet1.row(count).replace movements.row(i)
              else
                sheet1.row(count).replace movements.row(i)
              end
            end
          end
        end
        count = count + 1
      end

      export_file_path = [Rails.root, "public", "cartola_correcta.xls"].join("/")
      book.write export_file_path
      send_file export_file_path, :content_type => "application/vnd.ms-excel", :disposition => 'inline'

      return render json: {status: true}
    rescue Exception => e
      puts e
      return render json: {status: false}
    end
  end

  # Summary: Creación de primera cartola
  #
  # Params: Archivo a precesar
  # Return: status
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