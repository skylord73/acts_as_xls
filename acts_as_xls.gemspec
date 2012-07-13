# Provide a simple gemspec so you can easily use your enginex
# project in your rails apps through git.
require File.expand_path('../lib/acts_as_xls/version', __FILE__)

Gem::Specification.new do |s|
  s.authors        = ["Andrea Bignozzi"]
  s.email            = ["skylord73@gmail.com"]
  s.description   = %q{Extend Rails capabilities of importing and exporting excel files thanks to Spreadsheet gem}
  s.summary      = %q{Extend Rails capabilities of importing and exporting excel files}
  
  s.files             = Dir["{app,lib,config}/**/*"] + ["MIT-LICENSE", "Rakefile", "Gemfile", "README.rdoc", "CHANGELOG.md"]
  s.executables   = s.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  s.test_files      = s.files.grep(%r{^(test|spec|features)/})
  s.name            = "acts_as_xls"
  s.require_paths   = ["lib"]
  s.version         = ActsAsXls::VERSION
  
  s.add_dependency "rails", "~>3.0.14"
  s.add_dependency "spreadsheet"
  s.add_dependency "axlsx"

end
