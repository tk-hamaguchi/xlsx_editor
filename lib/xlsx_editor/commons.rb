module XlsxEditor
  module Commons
    def sheet_names_for(path)
      sheet_names = {}
      Zip::InputStream.open(path, 0) do |input|
        while (entry = input.get_next_entry)
          next unless entry.name == 'xl/workbook.xml'
          workbook = REXML::Document.new(input.read)
          workbook.elements['//sheets'].each do |sheet|
            sheet_names[sheet.elements['@name'].to_s.strip] = sprintf(
              "xl/worksheets/sheet%d.xml",
              sheet.elements['@r:id'].to_s.strip.match(/^rId(\d+)$/)[1]
            )
          end
        end
      end
      sheet_names
    end

    def shared_items_for(path)
      si = []
      Zip::InputStream.open(path, 0) do |input|
        while (entry = input.get_next_entry)
          next unless entry.name == 'xl/sharedStrings.xml'
          shared_strings = REXML::Document.new(input.read)
          shared_strings.elements['//sst'].each_with_index do |s,i|
            begin
              si[i] = s.elements['t'].text.strip.gsub(/　/, ' ').gsub(/\s+/, ' ')
            rescue
            end
          end
        end
      end
      si
    end
  end
end
