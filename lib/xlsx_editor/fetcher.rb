module XlsxEditor
  module Fetcher
    def fetch(path:, sheet_name:, cells:)
      result = {}

      sheet_names = sheet_names_for(path)
      si          = shared_items_for(path)

      Zip::InputStream.open(path, 0) do |input|
        while (entry = input.get_next_entry)
          next unless entry.name == sheet_names[sheet_name]
          begin
            work_sheet = REXML::Document.new(input.read)
            cells.each do |r|
              elm = work_sheet.elements["//c[@r=\"#{r}\"]"]
              result[r] = case elm.elements['@t'].to_s
                          when 's'
                            si[elm.elements['v'].text.to_i]
                          end
            end
          rescue
          end
        end
      end
      result
    end
  end
end
