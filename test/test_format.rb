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
# TestFormat -- Spreadsheet::ParseExcel -- 10.06.2003 -- hwyss@ywesee.com 

$: << File.expand_path("../lib", File.dirname(__FILE__))

require 'test/unit'
require 'parseexcel/format'

module Spreadsheet
	module ParseExcel
		class Format
			attr_writer :index
		end
	end
end
class StubFormatCell
	attr_accessor :numeric	
end

class TestFormat < Test::Unit::TestCase
	def setup
		@format = Spreadsheet::ParseExcel::Format.new
	end
	def test_cell_type
		cell = StubFormatCell.new
		assert_equal(:text, @format.cell_type(cell))
		cell.numeric = true
		assert_equal(:numeric, @format.cell_type(cell))
		@format.fmt_idx = 0x12
		assert_equal(:date, @format.cell_type(cell))
		@format.fmt_idx = 0
		assert_equal(:numeric, @format.cell_type(cell))
		@format.add_text_format(0x46, 'General')
		@format.fmt_idx = 0x46
		assert_equal(:numeric, @format.cell_type(cell))
	end
	def test_text_format
		assert_equal('foo', @format.text_format('foo'))
		assert_equal('foo', @format.text_format('foo', :_native_))
		assert_equal('foo', @format.text_format("\x00f\x00o\x00o", :ucs2))
	end
end
