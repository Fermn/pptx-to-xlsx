require 'sinatra'
require 'caxlsx'
require 'ruby_powerpoint'
require_relative 'lib/pptx_parser'

class App < Sinatra::Base
  get '/' do
    erb :index
  end

  post '/convert' do
    # Make sure file was uploaded
    if params['file']
      tempfile = params['file'][:tempfile]
      name = params['file'][:filename]
      unless tempfile && name
        return "File upload incomplete"
      end
    else 
      return "No file uploaded"
    end

    puts "Tempfile: #{tempfile.inspect}"
    puts "File path: #{tempfile.path}"
    # Parse PowerPoint file
    pptx_data = PPTXParser.parse(tempfile.path)

    # Create Excel file
    p = Axlsx::Package.new
    wb = p.workbook
    wb.add_worksheet(:name => "FINAL_PO") do |sheet|
      sheet.add_row [:Title, :Color_Ways, :Size_Card]
      pptx_data.each do |row|
        title = row[:title] || ''
        color_ways = row[:color_ways] || []
        size_card = row[:size_card] || ''

        if color_ways.empty?
          sheet.add_row [title, size_card, '']
        else
          color_ways.each do |color|
            sheet.add_row [title, color, size_card]
          end
        end
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
  trap('INT') do
    puts "Shutting down..."
    exit
  end
end
