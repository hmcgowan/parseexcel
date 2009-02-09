# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = %q{parseexcel}
  s.version = "0.5.3"
  s.platform = %q{universal-darwin-9}

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Hugh McGowan"]
  s.date = %q{2009-02-08}
  s.description = %q{Spreadsheet::ParseExcel - Get information from an Excel file.}
  s.email = %q{hugh_mcgowan@yahoo.com}
  s.extra_rdoc_files = ["README", "COPYING"]
  s.files = ["lib/parseexcel", "lib/parseexcel/format.rb", "lib/parseexcel/olestorage.rb", "lib/parseexcel/parseexcel.rb", "lib/parseexcel/parser.rb", "lib/parseexcel/workbook.rb", "lib/parseexcel/worksheet.rb", "lib/parseexcel.rb", "test/data", "test/data/annotation.xls", "test/data/bar.xls", "test/data/comment.5.0.xls", "test/data/comment.xls", "test/data/dates.xls", "test/data/float.5.0.xls", "test/data/float.xls", "test/data/foo.xls", "test/data/image.xls", "test/data/nil.xls", "test/data/umlaut.5.0.xls", "test/data/umlaut.biff8.xls", "test/data/uncompressed.str.xls", "test/suite.rb", "test/test_format.rb", "test/test_olestorage.rb", "test/test_parser.rb", "test/test_workbook.rb", "test/test_worksheet.rb", "README", "COPYING"]
  s.has_rdoc = true
  s.homepage = %q{http://github.com/hmcgowan/rasta}
  s.rdoc_options = ["--main", "README", "--inline-source", "--charset=UTF-8"]
  s.require_paths = ["lib"]
  s.rubyforge_project = %q{parseexcel}
  s.rubygems_version = %q{1.3.1}
  s.summary = %q{parseexcel}

  if s.respond_to? :specification_version then
    current_version = Gem::Specification::CURRENT_SPECIFICATION_VERSION
    s.specification_version = 2

    if Gem::Version.new(Gem::RubyGemsVersion) >= Gem::Version.new('1.2.0') then
    else
    end
  else
  end
end
