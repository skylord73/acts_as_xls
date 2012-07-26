require "action_controller"
require "action_view/template"

Mime::Type.register("application/vnd.ms-excel", :xls) unless Mime::Type.lookup_by_extension(:xls)
Mime::Type.register("application/vnd.openxmlformats", :xlsx) unless Mime::Type.lookup_by_extension(:xlsx)
#Mime::Type.register("application/pdf", :pdf) unless Mime::Type.lookup_by_extension(:pdf)

#Renderers try to render obj.to_xls
#then look for *.xls.maker template to render
ActionController::Renderers.add :xls do |obj, options|
  options ||= {}
  filename = options[:filename] || "file.xls"
  #Rails::logger.info("Renderer(xls) options=#{options.inspect}, obj=#{obj.inspect}")
  if obj.respond_to?(:to_xls)
    #No template, must render obj
    send_data obj.to_xls, :type => Mime::XLS, :filename => filename
  else
    send_file render_to_string(options[:template].to_s), :type => Mime::XLS, :disposition => "attachment; filename=#{filename}"
  end
  
end

#Renderers try to render obj.to_xlsx
#then look for *.xlsx.maker template to render
ActionController::Renderers.add :xlsx do |obj, options|
  options ||= {}
  filename = options[:filename] || "file.xlsx"
  Rails::logger.info("Renderer(xlsx) options=#{options.inspect}, obj=#{obj.inspect}")
  if obj.respond_to?(:to_xlsx)
    #No template, must render obj
    send_data obj.to_xlsx, :type => Mime::XLSX, :filename => filename
  else
    #Rails::logger.info("Renderer path=#{path}")
    send_file render_to_string(options[:template].to_s), :type => Mime::XLSX, :disposition => "attachment; filename=#{filename}"
  end
end


#~ ActionController::Renderers.add :pdf do |obj, options|
  #~ options ||= {}
  #~ filename = options[:filename] || "file.pdf"
  #~ #Rails::logger.info("Renderer(xlsx) options=#{options.inspect}, obj=#{obj.inspect}")
  #~ if obj.respond_to?(:to_pdf)
    #~ #No template, must render obj
    #~ send_data obj.to_xlsx, :type => Mime::PDF, :filename => filename
  #~ else
    #~ #Rails::logger.info("Renderer path=#{path}")
    #~ send_file render_to_string(options[:template]), :type => Mime::PDF, :disposition => "attachment; filename=#{filename}"
  #~ end
#~ end

module ActsAsXls
  #Render a template like *.[xls|xlsx].maker
  #the template is evaluated as a ruby expression.
  #you can use package (xlsx) or book (xls) inside the template to create an excel file 
  #template return tempfile path tha is rendered by the renderer.
  module MAKER
    
    def self.call(template)
      %(extend #{Proxy}; #{template.source}; render)
    end
    
    module Proxy
        def book
          Spreadsheet.client_encoding = 'UTF-8'
          @book ||= Spreadsheet::Workbook.new
        end
        
        def render_book
          @book.write @temp.path if @book
        end
                
        def package
          @package ||= Axlsx::Package.new
        end
        
        def render_package
          @package.serialize @temp.path if @package
        end
        
        
        #~ def pdf
          #~ @pdf ||= Prawn::Document.new(:page_size => "A4", :margin => [10.28,10,10,10])
        #~ end
        
        #Renders all methods starting form render_
        #this solution permits to extend che module adding new extesnions and renderers
        def render
          @temp = Tempfile.new("tmp")
          renderers = self.methods.select {|method| method.match(/^render_.*/)}
          Rails::logger.info("ActsAsXls::MAKER::Proxy renderers=#{renderers.inspect}")
          renderers.each {|r| eval(r)}
          @temp.path
        end
    end
    
  end
end


ActionView::Template.register_template_handler :maker, ActsAsXls::MAKER
