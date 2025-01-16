require 'ruby_powerpoint'

class PPTXParser
  # Parse a PowerPoint file for specific labels and their values
  # @param path [String] The file path of the PowerPoint file to parse
  # @return [Array<Hash>] An Array of hashes containing the parsed data or an empty array if any error occurs
  def self.parse(path)
    data = []
    begin
      deck = RubyPowerpoint::Presentation.new(path)
    rescue Errno::ENOENT => e
      # File not found error
      puts "Error: PowerPoint file not found at #{path}."
      return []
    rescue RubyPowerpoint::Error => e # Note: verify this error class name with the documentation for the gem
      # Handle RubyPowerpoint specific errors
      puts "Error parsing PowerPoint file: #{e.message}"
      return []
    rescue => e
      # Catch any other errors
      puts "An unexpected error occurred: #{e.message}"
      return []
    end

    deck.slides.each do |slide|
      slide_data = {}
      slide.content.each do |text|
        lower_text = text.downcase
        if lower_text.include?("title:")
          slide_data[:title] = text.split(':', 2).last.strip
        elsif lower_text.include?("size-card:")
          slide_data[:size_card] = text.split(':', 2).last.strip
        elsif lower_text.include?("color-ways:")
          slide_data[:color_ways] = text.split(':', 2).last.strip
        end
      end
      data << slide_data unless slide_data.empty?
    end
    data
  end
end
