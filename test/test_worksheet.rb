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
# TestWorksheet -- Spreadsheet::ParseExcel -- 10.06.2003 -- hwyss@ywesee.com 

$: << File.expand_path("../lib", File.dirname(__FILE__))

require 'test/unit'
require 'parseexcel/worksheet'
require 'parseexcel/parser'

module Spreadsheet
	module ParseExcel
		class Worksheet
			attr_accessor :cells
			attr_reader :min_col, :max_col, :min_row, :max_row, :row_heights
		end
	end
end
class StubWorksheetWorkbook
	attr_accessor :flg_1904
end

class TestWorksheet < Test::Unit::TestCase
	def setup
		@sheet = Spreadsheet::ParseExcel::Worksheet.new
	end
	def test_add_cell
		@sheet.add_cell(0,0,'foo')
		assert_equal([['foo']], @sheet.cells)
		assert_equal(0, @sheet.min_col)
		assert_equal(0, @sheet.max_col)
		assert_equal(0, @sheet.min_row)
		assert_equal(0, @sheet.max_row)
		@sheet.add_cell(1,2,'bar')
		assert_equal([['foo'],[nil,nil,'bar']], @sheet.cells)
		assert_equal(0, @sheet.min_col)
		assert_equal(2, @sheet.max_col)
		assert_equal(0, @sheet.min_row)
		assert_equal(1, @sheet.max_row)
	end
	def test_cell
		@sheet.cells = [['foo'],[nil,nil,'bar']]
		assert_equal('foo', @sheet.cell(0,0))
		assert_equal('bar', @sheet.cell(1,2))
		assert_instance_of(Spreadsheet::ParseExcel::Worksheet::Cell, @sheet.cell(1,0))
	end
	def test_row
		@sheet.cells = [['foo'],[nil,nil,'bar']]
		assert_equal(['foo'], @sheet.row(0))
		assert_equal([nil, nil, 'bar'], @sheet.row(1))
		assert_equal([], @sheet.row(2))
	end
	def test_set_dimensions1
		assert_equal(nil, @sheet.min_col)
		assert_equal(nil, @sheet.max_col)
		assert_equal(nil, @sheet.min_row)
		assert_equal(nil, @sheet.max_row)
		@sheet.set_dimensions(1,1)
		assert_equal(1, @sheet.min_col)
		assert_equal(1, @sheet.max_col)
		assert_equal(1, @sheet.min_row)
		assert_equal(1, @sheet.max_row)
		@sheet.set_dimensions(0,2)
		assert_equal(1, @sheet.min_col)
		assert_equal(2, @sheet.max_col)
		assert_equal(0, @sheet.min_row)
		assert_equal(1, @sheet.max_row)
	end
	def test_set_dimensions2
		assert_equal(nil, @sheet.min_col)
		assert_equal(nil, @sheet.max_col)
		assert_equal(nil, @sheet.min_row)
		assert_equal(nil, @sheet.max_row)
		@sheet.set_dimensions(1,1,2)
		assert_equal(1, @sheet.min_col)
		assert_equal(2, @sheet.max_col)
		assert_equal(1, @sheet.min_row)
		assert_equal(1, @sheet.max_row)
		@sheet.set_dimensions(0,2)
		assert_equal(1, @sheet.min_col)
		assert_equal(2, @sheet.max_col)
		assert_equal(0, @sheet.min_row)
		assert_equal(1, @sheet.max_row)
	end
	def test_set_row_height
		@sheet.set_row_height(2, 0)
		assert_equal([nil, nil, 0], @sheet.row_heights)
		@sheet.set_row_height(1, 4)
		assert_equal([nil, 4, 0], @sheet.row_heights)
	end
	def test_enumerable
		@sheet.cells = [[1],[2],[3],[4],[5]]
		result = @sheet.collect { |row| row.first }
		assert_equal([1,2,3,4,5], result)
		result = []
		@sheet.each(2) { |row|
			result << row.first
		}
		assert_equal([3,4,5], result)
	end
end
class TestCell < Test::Unit::TestCase
	def setup
		@cell = Spreadsheet::ParseExcel::Worksheet::Cell.new
	end
	def test_to_s
		@cell.value = 'foo'
		assert_equal('foo', @cell.to_s)
		@cell.value = 5
		assert_equal('5', @cell.to_s)
	end
	def test_to_i
		@cell.value = 'foo'
		assert_equal(0, @cell.to_i)
		@cell.value = 5
		assert_equal(5, @cell.to_i)
		@cell.value = '123'
		assert_equal(123, @cell.to_i)
	end
	def test_datetime
		@cell.value = 30025
		@cell.book = StubWorksheetWorkbook.new
		dt = @cell.datetime
		assert_equal(1982, dt.year)
		assert_equal(3, dt.month)
		assert_equal(15, dt.day)
		assert_equal(0, dt.hour)
		assert_equal(0, dt.min)
		assert_equal(0, dt.sec)
		assert_equal(0, dt.msec)
	end
	def test_datetime2
		@cell.value = nil
		@cell.book = StubWorksheetWorkbook.new
		assert_nothing_raised { 
			@cell.datetime
		}
	end
	def test_file_dates
		source = File.expand_path('data/dates.xls', File.dirname(__FILE__))
		parser = Spreadsheet::ParseExcel::Parser.new
		book = parser.parse(source)
		sheet = book.worksheet(0)
		expected = Date.new(1900,3,1)
		result = sheet.cell(0,0).date
		assert_equal(expected, result, "Expected #{expected} but was #{result}")
		expected = Date.new(1950,1,1)
		result = sheet.cell(1,0).date
		assert_equal(expected, result, "Expected #{expected} but was #{result}")
		expected = Date.new(1984,3,1)
		result = sheet.cell(2,0).date
		assert_equal(expected, result, "Expected #{expected} but was #{result}")
		expected = Date.new(2000,3,1)
		result = sheet.cell(3,0).date
		assert_equal(expected, result, "Expected #{expected} but was #{result}")
		expected = Date.new(2100,3,1)
		result = sheet.cell(4,0).date
		assert_equal(expected, result, "Expected #{expected} but was #{result}")
		expected = Date.new(2003,3,1)
		result = sheet.cell(5,0).date
		assert_equal(expected, result, "Expected #{expected} but was #{result}")
	end
end
