# frozen_string_literal: true

source "https://rubygems.org"

gem 'sinatra', '~> 4.1', '>= 4.1.1'
gem 'axlsx', '~> 2.0', '>= 2.0.1'
gem 'ruby_powerpoint', '~> 1.4', '>= 1.4.4'

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
end
