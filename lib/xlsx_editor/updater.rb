module XlsxEditor
  module Updater
    def update(path:, sheet_name:, cells:, output_file:)
      result = {}

      sheet_names = sheet_names_for(path)
      si          = shared_items_for(path)

      buffer = Zip::OutputStream.write_buffer(::StringIO.new('')) do |out|
        Zip::InputStream.open(f, 0) do |input|
          while (entry = input.get_next_entry)
            out.put_next_entry(entry.name)
            file_buf = input.read

            if entry.name == sheet_names[work_sheet_name]
              sheet = REXML::Document.new(file_buf)
              cells.each do |k, v|
                elm = sheet.elements["//c[@r=\"#{k}\"]"]
                if si.include?(v)
                  elm.add_attribute('t', 's')
                  unless v_elm = elm.elements['v']
                    v_elm = elm.add_element('v')
                  end
                  v_elm.text = si.index(v)
                else
                  elm.add_attribute('t', 'str')
                  unless v_elm = elm.elements['v']
                    v_elm = elm.add_element('v')
                  end
                  v_elm.text = v
                end
                print elm.to_s + "\n"
              end
              file_buf = sheet.to_s
            end

            out.write file_buf
          end
        end
      end
      File.open(output_file, 'wb') { |f| f.write(buffer.string) }
    end
  end
end
