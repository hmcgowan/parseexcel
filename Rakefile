$:.unshift('lib')
 
require 'rake/testtask'

begin
  require 'jeweler'
  Jeweler::Tasks.new do |s|
    s.name               = 'parseexcel'
    s.rubyforge_project  = 'parseexcel'
    s.platform           = Gem::Platform::CURRENT
    s.email              = 'hwyss@ywesee.com' 
    s.homepage           = "http://raa.ruby-lang.org/project/parseexcel/"
    s.summary            = "parseexcel"
    s.description        = "Spreadsheet::ParseExcel - Get information from an Excel file."
    s.authors            = ['Hannes Wyss']
    s.files              =  FileList[ "{lib,test}/**/*"]
    s.has_rdoc = true
    s.extra_rdoc_files = ["README", "COPYING"]
    s.rdoc_options = ["--main","README"]
  end
rescue LoadError
  puts "Jeweler not available. Install it with: sudo gem install technicalpickles-jeweler -s http://gems.github.com"
end

Rake::TestTask.new do |t|
  t.libs << "test"
  t.test_files = FileList['test/test*.rb']
  t.verbose = true
end


task :default => :test

