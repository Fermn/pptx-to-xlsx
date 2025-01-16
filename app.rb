require 'sinatra'
require 'axlsx'
require 'ruby_powerpoint'
require_relative 'lib/pptx_parser'
class App < Sinatra::Base
  get '/' do
    erb :index
  end

  post '/convert' do
    # Make sure file was uploaded
    unless params[:file] && (tmpfile = params[:file][:tempfile]) && (name = params[:file][:filename])
      return "No file uploaded"
    end

    # Parse PowerPoint file
    pptx_data = PPTXParser.parse(tmpfile.path)

    # Create Excel file
    p = Axlsx::Package.new
    wb = p.workbook
    wb.add_worksheet(:name => "Data") do |sheet|
      sheet.add_row[:Title, :Size_Card, :Color_Ways]
      pptx_data.each do |row|
        sheet.add_row row.values
      end
    end

    # Safe file to public folder
    excel_filename = "output_#{Time.now.to_i}.xlsx"
    p.serialize("public/#{excel_filename}")

    # redirect to results page
    redirect "/result?file=#{excel_filename}"
  end

  get '/result' do
    @file = params[:file]
    erb :result
  end

  run! if app_file == $0
end
