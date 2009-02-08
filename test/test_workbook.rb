#!/usr/bin/env ruby
#
#	Spreadsheet::ParseExcel -- Extract Data from an Excel File
#	Copyright (C) 2003 ywesee -- intellectual capital connected
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
#
#	ywesee - intellectual capital connected, Winterthurerstrasse 52, CH-8006 Zürich, Switzerland
#	hwyss@ywesee.com
#
# TestWorkbook -- Spreadsheet::ParseExcel -- 10.06.2003 -- hwyss@ywesee.com 

$: << File.expand_path("../lib", File.dirname(__FILE__))

require 'test/unit'
require 'parseexcel/workbook'

module Spreadsheet
	module ParseExcel
		class Workbook
			attr_accessor :worksheets, :formats
		end
	end
end

class TestWorkbook < Test::Unit::TestCase
	def setup
		@workbook = Spreadsheet::ParseExcel::Workbook.new
	end
	def test_format
		@workbook.format = :foo
		@workbook.formats = [:bar, :baz]
		assert_equal(:foo, @workbook.format)
		assert_equal(:bar, @workbook.format(0))
		assert_equal(:baz, @workbook.format(1))
	end
	def test_sheet_count
		assert_equal(0, @workbook.sheet_count)
		@workbook.worksheets = [:foo, :bar]
		assert_equal(2, @workbook.sheet_count)
		@workbook.worksheets = [:foo]
		assert_equal(1, @workbook.sheet_count)
	end
	def test_worksheet
		@workbook.worksheets = []
		sheet0 = @workbook.worksheet(0)
		assert_instance_of(Spreadsheet::ParseExcel::Worksheet, sheet0)
		assert_equal([sheet0], @workbook.worksheets)
		sheet1 = @workbook.worksheet(1)
		assert_instance_of(Spreadsheet::ParseExcel::Worksheet, sheet0)
		assert_not_equal(sheet0, sheet1)
		assert_equal([sheet0, sheet1], @workbook.worksheets)
		sheet2 = @workbook.worksheet(0)
		assert_instance_of(Spreadsheet::ParseExcel::Worksheet, sheet2)
		assert_equal(sheet0, sheet2)
		assert_equal([sheet0, sheet1], @workbook.worksheets)
	end
end
