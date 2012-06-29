module ActsAsXls
  class CustomError < StandardError
    def initialize(*args)
        @options = args.extract_options!
        super
    end
      
    def message
      I18n.t("#{self.class.name.gsub(/::/,'.')}", @options)
    end
  end 
  

end

