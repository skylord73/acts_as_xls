module ActsAsXls
  module Xls
    class << self
      def new_from_xls(base, *args)
        base.inspect
      end
      
      def to_xls(base_instance, *args, &block)
        #Rails::logger.info("Array#to_xls dentro")
        options = args.extract_options!
        
        if base_instance.kind_of?(ActiveRecord::Base)
            #Single ActiveRecord record
            return to_xls_from_activerecord(base_instance, *(args << options), &block) if base_instance.kind_of?(ActiveRecord::Base)
        elsif base_instance.kind_of?(Array) || base_instance.kind_of?(ActiveRecord::Relation)
            #Array of Array or Array of ActiveRecord
            return to_xls_from_activerecord(nil, *(args << options.merge!({:prepend => base_instance})), &block) if base_instance.first.kind_of?(Array)
            return to_xls_from_activerecord(base_instance, *(args << options), &block) if base_instance.first.kind_of?(ActiveRecord::Base)
        else
          "UnknownType"
        end
      end
      
      private
      
      
      #==Usage
      # $ Item.all.to_xlsLa classe restituisce un excel
      #
      #Return a binary string of excel with columns title (humanize columns name) and a row for each record
      #Columns name are used as method to render cell values:
      # column_name = name   => cell_value = @item.name
      # column_name = user.name   => cell_value = @item.user.name
      #
      #==Options
      # :only => [...]                                      #List of columns to render
      # :except => [...]                                  #List of columns not to render (must be in humanized form)
      # :header => true                                 #Add columns name as title
      # :header_columns => ["col1", "col2", ..]   #Set columns name
      # :human => true                                  #Humanize columns name using ActiveRecord Localization
      # :worksheet => "Foglio1"                      #worksheet name
      # :file => true                                     #Return a Filetmp instead of a StringIO
      # :prepend => [["Col 0, Row 0", "Col 1, Row 0"], ["Col 0, Row 1"]]
      #   Prepend columns and row passed to excel
      # 
      # block |column, value, row_index| ....
      def to_xls_from_activerecord(base_instance, *args, &block)
        options = args.extract_options!
        
        return nil if empty?(base_instance) && options[:prepend].blank?

        #Rails::logger.info("Array#to_xls_from_activerecord(1) dopo primo return")

        columns = []
        options.reverse_merge!(:header => true)
        options[:human] ||= true
        options[:worksheet] ||= "Foglio1"
        xls_report = StringIO.new
        book = Spreadsheet::Workbook.new
        sheet = book.create_worksheet
        sheet.name = options[:worksheet]
        
        sheet_index = 0
               
        unless options[:prepend].blank?
          options[:prepend].each do |array|
            sheet.row(sheet_index).concat(array)
            sheet_index += 1
          end
        end
        
        unless empty?(base_instance)
          
          base_instance = base_instance.to_a
          
          
          #Must check culumns name agains methids and removing the non working
          if options[:only]
            #columns = Array(options[:only]).map(&:to_sym)
            columns = Array(options[:only]).map { |column| validate(column,base_instance)}
            columns.compact!
          else
            columns = base_instance.first.class.column_names.map(&:to_sym) - Array(options[:except]).map(&:to_sym)
          end
          
          #Humanize columns name
          options[:header_columns] ||= columns.map do |column|
            #Se :human => false non umanizzo i nomi
            #Iconve risolve i problemi di codifica su caratteri strani tipo à che diventa \340....
            options[:human] == false ? column.to_s : Iconv.conv("UTF-8", "LATIN1", base_instance.first.class.human_attribute_name(column.to_s.gsub(/\./,"_")))
          end
        
          return nil if columns.empty? && options[:prepend].blank?
        
          #Add headers to excel
          if options[:header]
            sheet.row(sheet_index).concat(options[:header_columns])
            sheet_index += 1
          end
      
          base_instance.each_with_index do |obj, index|
            if block
              sheet.row(sheet_index).replace(columns.map { |column| block.call(column, get_column_value(obj, column), index) })
            else
              sheet.row(sheet_index).replace(columns.map { |column| get_column_value(obj, column) })
            end
            
            sheet_index += 1
          end
        end
        
        #Rails::logger.info("Spreadsheet::Workbook sheet.rows=#{sheet.rows.inspect}")
        #Rails::logger.info("Spreadsheet::Workbook book=#{book.inspect}")
        if options[:file] == true
          Rails::logger.info("Spreadsheet::Workbook file")
          tmp = Tempfile.new('tmp_xls')
          book.write(tmp.path)
          tmp
        else
          Rails::logger.info("Spreadsheet::Workbook stringIO")
          book.write(xls_report)
          xls_report.string
        end
      end
    
      #Rails::logger.info("Array#to_xls_from_activerecord obj.option=#{options.inspect}, self.inspect=#{self.inspect}")
      #Restituisce il valore della colonna comprese le tabelle collegate
      #Nel caso il valore della tabella collegata sia nullo restiuisce ""
      def get_column_value(obj, column) #:doc:
        if column.to_s.split(".").length > 1
          securize(obj, column) ? obj.instance_eval(column.to_s) : ""
        else
          obj.send(column)
        end
      end
  
      def securize(obj, column)
        collection = column.to_s.split(".")
        if collection.length > 1
          collection.pop
          collection = collection.join(".")
          !obj.instance_eval(collection).blank?
        else
          true
        end   
      end
      
      #Return true if base instance is empty or null
      def empty?(base_instance)
        return true if base_instance.nil? 
        return true if base_instance.blank?
        false
      end
      
      def validate(column,base_instance)
        base_instance.first.instance_eval("self." + column.to_s)
        column.to_sym
      rescue NoMethodError, RuntimeError
        nil
      end
      
      
    end
  end
end
