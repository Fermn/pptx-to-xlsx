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

  
end
