class Document
    def initialize(file_name)
        @file_name= file_name
        @file_path= File.expand_path("data/#{file_name}")
    end

    def load_xls_file(keys)
        require 'spreadsheet'
        doc = Spreadsheet.open(@file_path)
        sheet1 = doc.worksheet 0
        rows= []
        sheet1.each do |row|
          data= row.to_a
          if keys.nil?
            data.shift
            rows << data
            next
          end

          line= []
          keys.each do |key|
            line << (key.to_s=='' ? '' : data[key])
          end
          rows << line
        end
        return rows
    end

    def load_split_xls_file(keys)
        list= load_xls_file(keys)
        #take first element and then annd number sign as first element
        header= list.shift.unshift("№")
        header[2]= 'Came' if header[2].empty?
        rows= []
        ["А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ш", "Щ", "Э", "Ю", "Я"].each do |letter|
            row= [letter, header] + list.select{|person| person[0][0]==letter}
            rows << row
        end
        #here we make numbers for each human in different letter
        rows.each do |row|
            #skip first two items necause they are letter and heading of table
            row.each_with_index{|line, index|
                next if index<2
                line.unshift(index-1)
            }
        end
        return rows
    end

    def make_doc_table_page(docx, items)
        letter= items.shift
        docx.h2 "Register list: #{letter}"
        docx.hr
        docx.table items, border_size: 4 do
            cell_style rows[0], background: 'cccccc', bold: true
            cell_style cols[0], width: 500
            cell_style rows, height: 150
            cell_style cells, size: 16
        end
        docx.page
    end

    def make_questionary_part(docx, list)
        rows= []
        number=0
        list.each do |items|
            letter= items.shift
            titles= items.shift
            rows[number]= []
            items.each_slice(3) do |first, second, third|
                @items= [first, second, third]
                rows[number] << questionary_row_template
                #need this to split questionaries to 3 rows with 3 cells each (9 records)
                if rows[number].size == 3
                    number += 1
                    rows[number]= []
                end
            end
        end

        rows.each do |items|
            next if items.empty?
            docx.table items, border_size: 1 do
                cell_style cells, size: 20
            end
            docx.page
        end
    end

    private
=begin
  number => 0
  fio     => 1
  tel     => 17(home) 18(mobile)
  email   => 12
  address => 13(country) 14(city) 15(address)
  zipcode => 16
=end
      def questionary_row_template
          result= []
          @items.each do |item|
            data= Caracal::Core::Models::TableCellModel.new margins: { top: 100, bottom: 0, left: 10, right: 10 } do
              unless item.nil?
                p "Questionary: #{item[0]}", bold: true
                p "Full name: #{item[1]}"
                p "E-mail: #{item[12]}"
                p
                p "Phone: #{item[17]} #{item[18]}"
                p

                addr_empty= item[16].nil? and item[13].nil? and item[14].nil? and item[15].nil?
                p "Address: #{item[16]} #{item[13]}, #{item[14]} #{item[15]}" unless addr_empty
                p "Address:" if addr_empty
                p
                p
              end
            end
            result << data
          end
          result
      end
end
