#!/usr/bin/env ruby
# TestSuite -- spreadsheet -- 02.06.2003 -- hwyss@ywesee.com 

$: << File.expand_path(File.dirname(__FILE__))
$: << File.expand_path('../lib', File.dirname(__FILE__))

Dir.foreach(File.dirname(__FILE__)) { |file|
	require file if /^test_.*\.rb$/o.match(file)
}
