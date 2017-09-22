require_relative '../lib/document'
namespace :people do
    desc "Creates alphabetic DOCX document from XLS file"
	task :make_list do
        document= Document.new('base_2017_02_12.xls')
        list= document.load_split_xls_file([1, ''])

        require 'caracal'
        Caracal::Document.save 'data/people_alphabet_list.docx' do |docx|
            list.each do |items|
                document.make_doc_table_page(docx, items)
            end
        end
    end

    desc "Creates questionary for each participant basing on XLS file"
    task :make_questionary do
        document= Document.new('base_2017_02_12.xls')
        list= document.load_split_xls_file(nil)

        require 'caracal'
        Caracal::Document.save 'data/people_questionaries.docx' do |docx|
            document.make_questionary_part(docx, list)
        end
    end
end
