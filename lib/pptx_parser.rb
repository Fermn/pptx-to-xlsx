require 'ruby_powerpoint'

class PPTXParser
  # Parse a PowerPoint file for specific labels and their values
  # @param path [String] The file path of the PowerPoint file to parse
  # @return [Array<Hash>] An Array of hashes containing the parsed data or an empty array if any error occurs
  def self.parse(path)
    data = []
    begin
      deck = RubyPowerpoint::Presentation.new(path)
    rescue Errno::ENOENT
      puts "File not found at path: #{path}"
      return []
    rescue Errno::EACCES
      puts "Permission denied for file: #{path}"
      return []
    rescue => e # General catch-all for any other Ruby exceptions
      puts "Unexpected error occurred: #{e.class} - #{e.message}"
      return []
    end

    deck.slides.each do |slide|
      slide_data = {}
      slide.content.each do |text|
        lower_text = text.downcase
        if lower_text.include?("title:")
          slide_data[:title] = text.split(':', 2).last.strip.downcase
        elsif lower_text.include?("size-card:")
          slide_data[:size_card] = text.split(':', 2).last.strip
        elsif lower_text.include?("color-ways:")
          color_ways = text.split(':', 2).last.strip.split(', ').map(&:downcase)
          slide_data[:color_ways] = color_ways
        end
      end
      data << slide_data unless slide_data.empty?
    end
    data
  end
end
