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
# TestOLEReader -- Spreadsheet::ParseExcel -- 05.06.2003 -- hwyss@ywesee.com 

$: << File.expand_path("../lib", File.dirname(__FILE__))

require 'test/unit'
require 'parseexcel/olestorage'

module OLE
class Storage
	public :get_header 
	class Header 
		attr_writer :bbd_info
	end
end
end

class TestOLEStorageClass < Test::Unit::TestCase
	def test_is_normal_block
		assert(OLE::Storage.is_normal_block?(0))
		assert(OLE::Storage.is_normal_block?(0xFF))
		assert(!OLE::Storage.is_normal_block?(0xFFFFFFFC))
	end
	def test_invalid_header
		filename = File.expand_path('data/nil.xls', File.dirname(__FILE__))
		file = File.open(filename)
		assert_raises(OLE::UnknownFormatError) {
			h = OLE::Storage::Header.new(file)
		}
		file.close
	end
	def test_asc2ucs
		expected = "W\000o\000r\000k\000b\000o\000o\000k\000"
		assert_equal(expected, OLE.asc2ucs('Workbook'))
		expected = "R\000o\000o\000t\000 \000E\000n\000t\000r\000y\000"
		assert_equal(expected, OLE.asc2ucs('Root Entry'))
	end
end
class TestOLEDateTime < Test::Unit::TestCase
	def test_year_days
		assert_equal(perl_year_days(2000), OLE::DateTime.year_days(2000))
		assert_equal(perl_year_days(1999), OLE::DateTime.year_days(1999))
		assert_equal(perl_year_days(1900), OLE::DateTime.year_days(1900))
		assert_equal(perl_year_days(1980), OLE::DateTime.year_days(1980))
		assert_equal(366, OLE::DateTime.year_days(2000))
		assert_equal(365, OLE::DateTime.year_days(1999))
		assert_equal(365, OLE::DateTime.year_days(1900))
		assert_equal(366, OLE::DateTime.year_days(1980))
	end
	def test_month_days
		assert_equal(31, OLE::DateTime.month_days(8,1975))
		assert_equal(30, OLE::DateTime.month_days(6,2003))
		assert_equal(29, OLE::DateTime.month_days(2,2000))
		assert_equal(29, OLE::DateTime.month_days(2,1980))
		assert_equal(28, OLE::DateTime.month_days(2,1900))
	end
	def test_parse
		# how?
	end
	def test_date
		datetime = OLE::DateTime.new(2002,4,19)
		assert_equal(Date.new(2002, 4, 19), datetime.date)
	end

	# helper methods
	def perl_year_days(year)
		perl_leap_year?(year) ? 366 : 365
	end
	def perl_leap_year?(iYear)
		(((iYear % 4)==0) && ((iYear % 100).nonzero? || (iYear % 400)==0))
	end
end
class TestOLEStorage < Test::Unit::TestCase
	def setup
		@datadir = File.expand_path('data', File.dirname(__FILE__))
		@filename = File.expand_path('foo.xls', @datadir)
		@ole = OLE::Storage.new(@filename)
	end
	def test_get_header
		header = @ole.get_header
		assert_equal(1, header.bdb_count) 
		assert_equal(512, header.big_block_size) 
		assert_equal(0, header.extra_bbd_count) 
		assert_equal(4294967294, header.extra_bbd_start) 
		assert_equal(11, header.root_start)
		assert_equal(1, header.sbd_count) 
		assert_equal(2, header.sbd_start) 
		assert_equal(64, header.small_block_size) 
	end
	def test_search_pps
		expected = [
			@ole.header.get_nth_pps(0),
			@ole.header.get_nth_pps(1),
		]
		result = @ole.search_pps([
			OLE.asc2ucs('Root Entry'), 
			OLE.asc2ucs('Workbook'),
		])
		assert_equal(expected, result)
		lowercase = [
			OLE.asc2ucs('root entry'), 
			OLE.asc2ucs('workbook'),
		]
		assert_equal([], @ole.search_pps(lowercase))
		result = @ole.search_pps(lowercase, true)
		assert_equal(expected, result)
	end
  def test_unknown_format
    assert_raises(OLE::UnknownFormatError) { 
      OLE::Storage.new(StringIO.new('12345678'))
    }
  end
end
class TestOLEStorageHeader < Test::Unit::TestCase
	def setup
		@filename = File.expand_path('data/foo.xls', File.dirname(__FILE__))
		@file = OLE::Storage.new(@filename)
		@header = @file.get_header
	end
	def test_get_next_block_no
		@header.bbd_info = {
			1	=>	2,
			3	=>	5,
		}
		assert_equal(2, @header.get_next_block_no(1))
		assert_equal(3, @header.get_next_block_no(2))
		assert_equal(5, @header.get_next_block_no(3))
	end
	def test_get_nth_block_no
		@header.bbd_info = {
			1	=>	2,
			3	=>	5,
		}
		assert_equal(2, @header.get_nth_block_no(1,1))
		assert_equal(3, @header.get_nth_block_no(1,2))
		assert_equal(5, @header.get_nth_block_no(1,3))
		assert_equal(3, @header.get_nth_block_no(2,1))
		assert_equal(5, @header.get_nth_block_no(2,2))
		assert_equal(5, @header.get_nth_block_no(3,1))
	end
	def test_get_nth_pps
		root = @header.get_nth_pps(0)
		assert_instance_of(OLE::Storage::PPS::Root, root)
		file = @header.get_nth_pps(1)
		assert_instance_of(OLE::Storage::PPS::File, file)
		assert_equal(1, root.dir_pps)
		min_1, = [-1].pack('V').unpack('V')
		assert_equal(min_1, root.next_pps)
		assert_equal(min_1, root.prev_pps)
		assert_equal(3904, root.data.size)
		assert_equal(3335, file.data.size)
	end
	def test_sb_start
		assert_equal(3, @header.sb_start)
	end
	def test_sb_size
		assert_equal(3904, @header.sb_size)
	end
end
