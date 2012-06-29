# Provide a simple gemspec so you can easily use your enginex
# project in your rails apps through git.
require File.expand_path('../lib/acts_as_xls/version', __FILE__)

Gem::Specification.new do |s|
  s.name = "acts_as_xls"
  s.summary = "Extend Rails capabilities of importing and exporting excel files"
  s.description = "Extend Rails capabilities of importing and exporting excel files thanks to Spreadsheet gem"
  s.files = Dir["{app,lib,config}/**/*"] + ["MIT-LICENSE", "Rakefile", "Gemfile", "README.rdoc", "CHANGELOG.md"]
  s.version = ActsAsXls::VERSION

end
