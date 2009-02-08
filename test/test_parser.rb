#!/usr/bin/env ruby
# TestParser -- Spreadsheet -- 02.06.2003 -- hwyss@ywesee.com 

$: << File.expand_path("../lib", File.dirname(__FILE__))

require 'test/unit'
require 'parseexcel/parser'

module Spreadsheet
	module ParseExcel
		class Parser
			attr_reader :proc_table
			attr_accessor :workbook, :current_sheet, :buff
			attr_writer :bigendian, :prev_pos, :prev_info, :prev_cond
		end
		class Workbook
			attr_reader :pkg_strs
		end
	end
end
class StubParserWorksheet
	attr_accessor :header, :footer, :row_height, :dimensions, :default_row_height
	attr_accessor :paper, :scale, :page_start, :fit_width, :fit_height
	attr_accessor :resolution, :v_resolution
	attr_accessor :header_margin, :footer_margin, :copies, :left_to_right
	attr_accessor :no_pls, :no_color, :draft, :notes, :no_orient, :use_page
	attr_accessor :landscape, :name
	def initialize
		@cells = []
	end
	def add_cell(x, y, cell)
		(@cells[x] ||= [])[y] = cell
	end
	def cell(x, y)
		(@cells[x] ||= [])[y]
	end
	def set_row_height(row, hght)
		@row_height = hght
	end
	def set_dimensions(row, scol, ecol)
		@dimensions = [row, scol, ecol]
	end
end
class StubParserWorkbook 
	attr_accessor :header, :footer, :worksheets, :text_format
	attr_accessor :biffversion, :version, :flg_1904
	attr_writer :format, :pkg_strs
	def add_cell_format(format)
		@format = format
	end
	def add_text_format(idx, fmt)
		@text_format = fmt
	end
	def worksheet(index)
		@worksheets[index] ||= StubParserWorksheet.new
	end
	def format(index=nil)
		@format
	end
	def sheet_count
		(@worksheets.respond_to? :size) ? @worksheets.size : 0
	end
	def pkg_str(idx)
		@pkg_strs[idx]
	end
end
class StubParserFormat
	attr_writer :cell_type
	def text_format(txt, code=nil)
		txt
	end
	def cell_type(foo)
		@cell_type
	end
end

class TestParser < Test::Unit::TestCase
	def setup
		@source = File.expand_path('data/foo.xls', File.dirname(__FILE__))
		@book = StubParserWorkbook.new
		@sheet = StubParserWorksheet.new
		@format = StubParserFormat.new
		@book.worksheets = [@sheet]
		@book.format = @format
		@parser = Spreadsheet::ParseExcel::Parser.new
		@parser.workbook = @book
		@parser.current_sheet = @sheet
		@parser.private_methods.each { |m| 
			@parser.class.instance_eval("public :#{m}")
		}
		@bin_header = "\x03\x00\x02\x00\x01\x00"
	end
	def test_initialize
		@parser.proc_table.each { |key, val|
			assert_instance_of(Method, val)
			assert_equal(3, val.arity)
		}
		bigendian = ([2].pack('L').unpack('H8').first !=	'02000000')
		assert_equal(bigendian, @parser.bigendian)
		parser = Spreadsheet::ParseExcel::Parser.new(:bigendian => false)
		assert_equal(false, parser.bigendian)
		parser = Spreadsheet::ParseExcel::Parser.new(:bigendian => true)
		assert_equal(true, parser.bigendian)
	end
	def test_blank
		@parser.blank(nil, nil, @bin_header)
		cell = @sheet.cell(3,2)
		assert_instance_of(Spreadsheet::ParseExcel::Worksheet::Cell, cell)
		assert_equal(:blank, cell.kind)
	end
	def test_bof
		bof = [0x500, 0x5].pack("v2")
		@parser.bof(nil, nil, bof)
		assert_nil(@parser.current_sheet)
		assert_equal(0x08, @book.biffversion)
		bof = [0x000, 0x3].pack("v2")
		@parser.bof(0x02*0x100, nil, bof)
		assert_instance_of(StubParserWorksheet, @parser.current_sheet)
	end
	def test_bool_err1
		@parser.bool_err(nil, nil, @bin_header + "\x01\x00")
		cell = @sheet.cell(3,2)
		assert_equal(:bool_error, cell.kind)
		assert_equal("TRUE", cell.value)
	end
	def test_bool_err2
		@parser.bool_err(nil, nil, @bin_header + "\x00\x00")
		cell = @sheet.cell(3,2)
		assert_equal("FALSE", cell.value)
		assert_equal(:bool_error, cell.kind)
	end
	def test_bool_err3
		@parser.bool_err(nil, nil, @bin_header + "\x01\x01")
		cell = @sheet.cell(3,2)
		assert_equal(:bool_error, cell.kind)
		assert_equal("#ERR", cell.value)
	end
	def test_bool_err4
		@parser.bool_err(nil, nil, @bin_header + "\x00\x01")
		cell = @sheet.cell(3,2)
		assert_equal(:bool_error, cell.kind)
		assert_equal("#NULL!", cell.value)
	end
	def test_cell_factory
		@format.cell_type = :text
		params = {
			:kind				=>	:blank,
			:value			=>	'',
			:format_no	=>	1,
			:numeric		=>	false,
		}
		cell = @parser.cell_factory(1,1,params)
		assert_equal('', cell.value)
		assert_equal(false, cell.numeric)
		assert_nil(cell.code)
		assert_equal(@book, cell.book)
		assert_equal(@format, cell.format)
		assert_equal(1, cell.format_no)
		assert_equal(:text, cell.type)
	end
	def test_conv_biff8_data1 # compressed
		biff8_data = [
			"\3\0\0foo"
		].join
		expected = [
			[ "f\0o\0o\0", false, "", "" ],
			6, 3, 3,
		]
		assert_equal(expected, @parser.conv_biff8_data(biff8_data))
	end
	def test_conv_biff8_data2 # short string
		biff8_data = [
			"\x04\x00\x00foo"
		].join
		expected = [
			[ nil, false, nil, nil ],
			7, 3, 4,
		]
		assert_equal(expected, @parser.conv_biff8_data(biff8_data))
	end
	def test_conv_biff8_data3 # uncompressed
		biff8_data = [
			"\3\0\1\0f\0o\0o"
		].join
		expected = [
			[ "\0f\0o\0o" , true, "", "" ],
			9, 3, 6,
		]
		assert_equal(expected, @parser.conv_biff8_data(biff8_data))
	end
	def test_conv_biff8_data4	# ext
		biff8_data = [
			"\x03\x00\x04\x01\x00\x00\x00foobar"
		].join
		expected = [
			[ "f\0o\0o\0" , false, "bar", "b" ],
			11, 7, 3,
		]
		assert_equal(expected, @parser.conv_biff8_data(biff8_data))
	end
	def test_conv_biff8_data5	# rich
		biff8_data = [
			"\x03\x00\x08\x01\x00foobaar"
		].join
		expected = [
			[ "f\0o\0o\0" , false, "", "" ],
			12, 5, 3,
		]
		assert_equal(expected, @parser.conv_biff8_data(biff8_data))
	end
	def test_conv_biff8_data6 # ext && rich
		biff8_data = [
			"\x03\x00\x0C\x01\x00\x01\x00\x00\x00foobaarbaz"
		].join
		expected = [
			[ "f\0o\0o\0" , false, "baz", "b" ],
			17, 9, 3,
		]
		assert_equal(expected, @parser.conv_biff8_data(biff8_data))
	end
	def test_conv_biff8_string1
		str = "\010\000\000DD/MM/YY"
		assert_equal("D\0D\0/\0M\0M\0/\0Y\0Y\0", @parser.conv_biff8_string(str))
	end
	def test_conv_dval
		@parser.bigendian = true
		assert_equal(0.0, @parser.conv_dval("\x00"))
		last_bit = "\x00\x00\x00\x00\x00\x00\x00\x01"
		frst_bit = "\x01\x00\x00\x00\x00\x00\x00\x00"
		assert_in_delta(4.940656458412465e-324, @parser.conv_dval(last_bit), 1e-300)
		assert_in_delta(7.291122019556397e-304, @parser.conv_dval(frst_bit), 1e-300)
		@parser.bigendian = false
		assert_in_delta(7.291122019556397e-304, @parser.conv_dval(last_bit), 1e-300)
		assert_in_delta(4.940656458412465e-324, @parser.conv_dval(frst_bit), 1e-300)
	end
	def test_decode_bool_err
		assert_equal("FALSE", @parser.decode_bool_err(0x00))
		assert_equal("TRUE", @parser.decode_bool_err(0x01))
		assert_equal("#NULL!", @parser.decode_bool_err(0x00, true))
		assert_equal("#DIV/0!", @parser.decode_bool_err(0x07, true))
		assert_equal("#VALUE!", @parser.decode_bool_err(0x0F, true))
		assert_equal("#REF!", @parser.decode_bool_err(0x17, true))
		assert_equal("#NAME?", @parser.decode_bool_err(0x1D, true))
		assert_equal("#NUM!", @parser.decode_bool_err(0x24, true))
		assert_equal("#N/A!", @parser.decode_bool_err(0x2A, true))
		assert_equal("#ERR", @parser.decode_bool_err(0x02, true))
	end
	def test_default_row_height
		@parser.default_row_height(nil,nil,"\x00\x00\x28\x00")
		assert_equal(2.0, @sheet.default_row_height)
	end
	def test_flg_1904
		@parser.flg_1904(nil, nil, "\x00\x00")
		assert_nil(@book.flg_1904)
		@parser.flg_1904(nil, nil, "\x01\x00")
		assert(@book.flg_1904)
	end
	def test_footer1
		@book.biffversion = 0x08
		@parser.footer(nil, nil, "\x03\x00\x00\x00")
		assert_equal(nil, @sheet.footer)
		@parser.footer(nil, nil, "\x03foo")
		assert_equal('foo', @sheet.footer)
		@parser.footer(nil, nil, "\x03foobarbaz")
		assert_equal('foo', @sheet.footer)
	end
	def test_footer2
		@book.biffversion = 0x18
		@parser.footer(nil, nil, "\x01\x00\x00\x00")
		assert_equal(nil, @sheet.footer)
		@parser.footer(nil, nil, "\x03\x00\x00foo")
		assert_equal("f\0o\0o\0", @sheet.footer)
		@parser.footer(nil, nil, "\x03\x00\x00foobarbaz")
		assert_equal("f\0o\0o\0", @sheet.footer)
	end
	def test_format
		@book.biffversion = 0x02
		fmt = "\x00\x00\x03foo"
		@parser.format(nil, nil, fmt)
		assert_equal("foo", @book.text_format)
		@book.biffversion = 0x18
		fmt = "\x00\x00\x03\000\000foo"
		@parser.format(nil, nil, fmt)
		assert_equal("f\0o\0o\0", @book.text_format)
		fmt = "\0\0\6\0\1f\0o\0o\0"
		@parser.format(nil, nil, fmt)
		assert_equal("f\0o\0o\0", @book.text_format)
	end
	def test_header1
		@book.biffversion = 0x08
		@parser.header(nil, nil, "\x03\x00\x00\x00")
		assert_equal(nil, @sheet.header)
		@parser.header(nil, nil, "\x03foo")
		assert_equal('foo', @sheet.header)
		@parser.header(nil, nil, "\x03foobarbaz")
		assert_equal('foo', @sheet.header)
	end
	def test_header2
		@book.biffversion = 0x18
		@parser.header(nil, nil, "\x01\x00\x00\x00")
		assert_equal(nil, @sheet.header)
		@parser.header(nil, nil, "\x03\x00\x00foo")
		assert_equal("f\0o\0o\0", @sheet.header)
		@parser.header(nil, nil, "\x03\x00\x00foobarbaz")
		assert_equal("f\0o\0o\0", @sheet.header)
	end
	def test_integer
		@parser.integer(nil, nil, @bin_header + "x\x1A\x00")
		cell = @sheet.cell(3,2)
		assert_equal(:integer, cell.kind)
		assert_equal(26, cell.value)
	end
	def test_label1
		@parser.label(nil, nil, @bin_header+'abfoo')
		cell = @sheet.cell(3,2)
		assert_equal(:label, cell.kind)
		assert_equal('foo', cell.value)
		assert_equal(:_native_, cell.code)
	end
	def test_label2
		@book.biffversion = 0x18
		@parser.label(nil, nil, @bin_header+"\x03\x003\x00f\x00o\x00o")
		cell = @sheet.cell(3,2)
		assert_equal(:label, cell.kind)
		assert_equal("\0f\0o\0o", cell.value)
		assert_equal(:ucs2, cell.code)
	end
	def test_label_sst
		foo = Spreadsheet::ParseExcel::Worksheet::PkgString.new("foo", nil, 0, 0)
		@book.pkg_strs = [nil, nil, nil, nil, foo]
		label = [
			"\x01\x00",
			"\x02\x00",
			"\x03\x00",
			"\x04\x00\x00\x00",
		].join
		@parser.label_sst(nil, nil, label)
		cell = @sheet.cell(1,2)
		assert_equal("foo", cell.value)
	end
	def test_mul_rk
		mul_rk = [
			"\x01\x00",
			"\x02\x00",
			"\x01\x00\x00\x00\xF0\x3F",
			"\x01\x00\x01\x00\xF0\x3F",
			"\x01\x00\x46\x56\x4B\x00",
			"\x04\x00",
		].join
		@parser.mul_rk(nil, nil, mul_rk)
		cell1 = @sheet.cell(1,2)
		assert_equal(:mul_rk, cell1.kind)
		assert_equal(1.0, cell1.value)
		cell2 = @sheet.cell(1,3)
		assert_equal(:mul_rk, cell2.kind)
		assert_equal(0.01, cell2.value)
		cell3 = @sheet.cell(1,4)
		assert_equal(:mul_rk, cell3.kind)
		assert_equal(1234321, cell3.value)
	end
	def test_mul_blank
		blanks = [
			"\x02\x00",
			"\x01\x00",
			"\x01\x00",
			"\x02\x00",
			"\x03\x00",
			"\x04\x00",
			"\x04\x00",
		].join
		@parser.mul_blank(nil, nil, blanks)
		cell1 = @sheet.cell(2,1)
		assert_equal(:mul_blank, cell1.kind)
		cell2 = @sheet.cell(2,2)
		assert_equal(:mul_blank, cell2.kind)
		cell3 = @sheet.cell(2,3)
		assert_equal(:mul_blank, cell3.kind)
		cell4 = @sheet.cell(2,4)
		assert_equal(:mul_blank, cell4.kind)
	end
	def test_number
		@parser.number(nil, nil, @bin_header + "\x00\x00\x00\x00\x00\x00\x00\x00")
		cell = @sheet.cell(3,2)
		assert_equal(:number, cell.kind)
		assert_equal(0.0, cell.value)
	end
	def test_rk
		rk = "\x01\x00\x02\x00\x04\x00\x46\x56\x4B\x00"
		@parser.rk(nil, nil, rk)
		cell = @sheet.cell(1,2)
		assert_equal(:rk, cell.kind)
		assert_equal(1234321, cell.value)
	end
	def test_row
		row = [
			"\x01\x00", #row
			"\x01\x00", #scol
			"\x04\x00", #ecol
			"\x14\x00", #hght
			"\x00\x00", #nil
			"\x00\x00", #nil
			"\x00\x20", #gr
			"\x00\x00", #xf
		]
		@parser.row(nil, nil, row.join)
		assert_equal([1,1,3], @sheet.dimensions)
		assert_equal(1.0, @sheet.row_height)
	end
	def test_rstring
		rstr = "\x01\x00\x01\x00\x04\x00\x03\x00foo"
		@parser.rstring(nil, nil, rstr)
		cell = @sheet.cell(1,1)
		assert_equal("foo", cell.value)
		assert_nil(cell.rich)
		rstr = "\x01\x00\x02\x00\x04\x00\x03\x00foorich"
		@parser.rstring(nil, nil, rstr)
		cell = @sheet.cell(1,2)
		assert_equal("foo", cell.value)
		assert_equal("rich", cell.rich)
	end
	def test_setup
		setup = [
			"\x01\x00",
			"\x02\x00",
			"\x03\x00",
			"\x04\x00",
			"\x05\x00",
			"\x06\x00",
			"\x07\x00",
			"\x08\x00",
			"\x00\x00\x00\x00\x00\x00\xF0\x3F",
			"\x00\x00\x00\x00\x00\x00\xF1\x3F",
			"\x0B\x00",
		].join
		@parser.setup(nil, nil, setup)
		assert_equal(1, @sheet.paper)
		assert_equal(2, @sheet.scale)
		assert_equal(3, @sheet.page_start)
		assert_equal(4, @sheet.fit_width)
		assert_equal(5, @sheet.fit_height)
		assert_equal(7, @sheet.resolution)
		assert_equal(8, @sheet.v_resolution)
		assert_equal(2.54, @sheet.header_margin)
		assert_equal(2.69875, @sheet.footer_margin)
		assert(!@sheet.left_to_right, "left_to_right")
		assert(@sheet.landscape, "landscape")
		assert(@sheet.no_pls, "no_pls")
		assert(!@sheet.no_color, "no_color")
		assert(!@sheet.draft, "draft")
		assert(!@sheet.notes, "notes")
		assert(!@sheet.no_orient, "no_orient")
		assert(!@sheet.use_page, "use_page")
	end
	def test_sst
		@parser.sst(nil, nil, '12345678foo')
		assert_equal('foo', @parser.buff)
	end
	def test_string1
		@parser.prev_pos = [1,2,3]
		@parser.string(nil, nil, "\x06foobar")
		cell = @sheet.cell(1,2)
		assert_equal(:string, cell.kind)
		assert_equal("foobar", cell.value)
	end
	def test_string2 # BIFF5
		@parser.prev_pos = [1,2,3]
		@book.biffversion = 0x08
		@parser.string(nil, nil, "\x00\x06foobar")
		cell = @sheet.cell(1,2)
		assert_equal(:string, cell.kind)
		assert_equal("foobar", cell.value)
	end
	def test_string3 # BIFF5
		@parser.prev_pos = [1,2,3]
		@book.biffversion = 0x18
		@parser.string(nil, nil, "\x06\x00\x00foobar")
		cell = @sheet.cell(1,2)
		assert_equal(:string, cell.kind)
		assert_equal("f\0o\0o\0b\0a\0r\0", cell.value)
		assert_equal("foobar", cell.to_s('latin1'))
	end
	def test_str_wk1
		@parser.str_wk('foo')
		assert_equal('foo', @parser.buff)
	end
	def test_str_wk2
		@parser.str_wk('foo', true)
		assert_equal('foo', @parser.buff)
	end
	def test_str_wk3
		@parser.buff = 'f'
		@parser.str_wk('foo', true)
		assert_equal('foo', @parser.buff)
	end
	def test_str_wk4
		@parser.buff = 'f'
		@parser.prev_cond = true
		@parser.prev_info = [0,1]
		@parser.str_wk("oo", true)
		assert_equal('foo', @parser.buff)
	end
	def test_str_wk5
		@parser.buff = 'f'
		@parser.prev_cond = true
		@parser.prev_info = [1,2]
		@parser.str_wk("\x01oo", true)
		assert_equal('foo', @parser.buff)
	end
	def test_str_wk6
		@parser.buff = 'f'
		@parser.prev_cond = true
		@parser.prev_info = [1,2]
		@parser.str_wk("\x01oo", true)
		assert_equal('foo', @parser.buff)
	end
	def test_str_wk7
		@parser.buff = "fo"
		@parser.prev_cond = 1
		@parser.prev_info = [1,2]
		@parser.str_wk("\x01o", true)
		assert_equal("foo", @parser.buff)
	end
	def test_str_wk8
		@parser.buff = "\x00fo"
		@parser.prev_cond = false
		@parser.prev_info = [1,3]
		@parser.str_wk("\x01o\x00", true)
		assert_equal("\x00f\x00o\x00o\x00", @parser.buff)
	end
	def test_unpack_rk_rec
		rk_rec = "\030\000\001\330\200@"
		result = @parser.unpack_rk_rec(rk_rec)
		expected = [25, 5.39]
		rk_rec = "\017\000o\010\000\000"	
		rk_rec = "\031\000o\010\000\000"
		result = @parser.unpack_rk_rec(rk_rec)
		expected = [25, 5.39]
		#assert_equal(expected, result)
		rk_rec = "\x01\x00\x00\x00\xF0\x3F"
		result = @parser.unpack_rk_rec(rk_rec)
		expected = [1, 1.0]
		assert_equal(expected, result)
		rk_rec = "\x01\x00\x01\x00\xF0\x3F"
		result = @parser.unpack_rk_rec(rk_rec)
		expected = [1, 0.01]
		assert_equal(expected, result)
		rk_rec = "\x01\x00\x46\x56\x4B\x00"
		result = @parser.unpack_rk_rec(rk_rec)
		expected = [1, 1234321]
		assert_equal(expected, result)
		rk_rec = "\x01\x00\x47\x56\x4B\x00"
		result = @parser.unpack_rk_rec(rk_rec)
		expected = [1, 12343.21]
		assert_equal(expected, result)
	end
	def test_xf
		@book.biffversion = 0x18
		fmt = "\006\000\244\000\001\000 \000\000\000\000\000\010\004\010\004\002\000\t\004"
		format = @parser.xf(nil, nil, fmt)
		assert_equal(6, format.font_no)
		assert_equal(164, format.fmt_idx)
		assert_not_nil(format.lock)
		assert_equal(false, format.hidden)
		assert_equal(false, format.style)
		assert_equal(false, format.key_123)
		assert_equal(0, format.align_h)
		assert_equal(false, format.wrap)
		assert_equal(2, format.align_v)
		assert_equal(false, format.just_last)
		assert_equal(0, format.rotate)
		assert_equal(0, format.indent)
		assert_equal(false, format.shrink)
		assert_equal(false, format.merge)
		assert_equal(0, format.read_dir)
		assert_equal([0,0,0,0], format.border_style)
		assert_equal([8,8,8,8], format.border_color)
		assert_equal([0,0,0], format.border_diag)
		assert_equal([0,9,8], format.fill)
	end
end
class TestParser2 < Test::Unit::TestCase
	def setup
		@parser = Spreadsheet::ParseExcel::Parser.new
	end
	def test_file_bar # Simple text-values
		source = File.expand_path('data/bar.xls', File.dirname(__FILE__))
		book = @parser.parse(source)
		assert_equal(1, book.sheet_count)
		sheet = book.worksheet(0)
		assert_equal('A1',sheet.cell(0,0).value)
		assert_equal('A2',sheet.cell(1,0).value)
		assert_equal('A3',sheet.cell(2,0).value)
		assert_equal('B1',sheet.cell(0,1).value)
		assert_equal('B2',sheet.cell(1,1).value)
		assert_equal('B3',sheet.cell(2,1).value)
		assert_equal('C1',sheet.cell(0,2).value)
		assert_equal('C2',sheet.cell(1,2).value)
		assert_equal('C3',sheet.cell(2,2).value)
	end
	def test_file_foo
		source = File.expand_path('data/foo.xls', File.dirname(__FILE__))
		book = @parser.parse(source)
		assert_equal(3, book.sheet_count)
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal("F\0o\0o\0",cell0.value)
		assert_equal(:text, cell0.type)
		assert_equal('UTF-16LE', cell0.encoding)
		assert_equal('Foo', cell0.to_s('latin1'))
		cell1 = sheet.cell(0,1)
		assert_equal(12,cell1.value)
		assert_equal(:numeric, cell1.type)
		cell2 = sheet.cell(1,0)
		assert_equal(27627,cell2.value)
		assert_equal(:date, cell2.type)
		# once formulas are implemented:
		# cell3 = sheet.cell(1,1)
		# assert_equal(27627,cell3.value)
		# assert_equal(:numeric, cell3.type)
	end
	def test_file_float
		source = File.expand_path('data/float.xls', File.dirname(__FILE__))
		book = @parser.parse(source)
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal(5.39,cell0.value)
	end
	def test_file_image
		source = File.expand_path('data/image.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
		cell0 = sheet.cell(48,0)
		assert_equal('TEST', cell0.to_s('latin1'))
	end
	def test_file_umlaut__5_0
		source = File.expand_path('data/umlaut.5.0.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal('WINDOWS-1252', cell0.encoding)
		assert_equal('ä', cell0.to_s('latin1'))
	end
	def test_file_umlaut__biff8
		source = File.expand_path('data/umlaut.biff8.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal('UTF-16LE', cell0.encoding)
		assert_equal('ä', cell0.to_s('latin1'))
	end
	def test_file_uncompressed_str
		source = File.expand_path('data/uncompressed.str.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal('UTF-16LE', cell0.encoding)
		assert_equal('Aaaaa Aaaaaaaaa Aaaaaa', cell0.to_s('latin1'))
	end
	def test_file_comment
		source = File.expand_path('data/comment.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal('cellcontent', cell0.to_s('latin1'))
		assert_equal('cellcomment', cell0.annotation)
    assert_equal('HW', cell0.annotation.author)
		cell1 = sheet.cell(1,1)
		assert_equal('cellcontent', cell1.to_s('latin1'))
		assert_equal('annotation', cell1.annotation)
	end
	def test_file_comment__5_0
		source = File.expand_path('data/comment.5.0.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal('cellcontent', cell0.to_s('latin1'))
		assert_equal('cellcomment', cell0.annotation)
		cell1 = sheet.cell(1,1)
		assert_equal('cellcontent', cell1.to_s('latin1'))
		assert_equal('annotation', cell1.annotation)
	end
	def test_file_comment__ds
		source = File.expand_path('data/annotation.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
		cell0 = sheet.cell(0,0)
		assert_equal('hello', cell0.to_s('latin1'))
		ann = cell0.annotation
		assert_equal("david surmon:\nnow is the time for all good men to come to the aid of their country!", ann)
		assert_equal('F', ann.author)
		cell1 = sheet.cell(0,1)
		assert_equal('there', cell1.to_s('latin1'))
		ann = cell1.annotation
		assert_equal("david surmon:\nwhat should this comment be? Now what?", ann)
		assert_equal('F', ann.author)
		cell2 = sheet.cell(0,2)
		assert_equal('whos', cell2.to_s('latin1'))
		cell3 = sheet.cell(1,0)
		assert_equal('I', cell3.to_s('latin1'))
	end
	def test_file_several_sheets
		source = File.expand_path('data/annotation.xls', File.dirname(__FILE__))
		book = nil
		assert_nothing_raised { 
			book = @parser.parse(source)
		}
		sheet = book.worksheet(0)
    assert_equal('First Worksheet', sheet.name('latin1'))
		sheet = book.worksheet(1)
    assert_equal('Second Worksheet', sheet.name('latin1'))
		cell0 = sheet.cell(0,0)
		assert_equal('version', cell0.to_s('latin1'))
		cell1 = sheet.cell(1,0)
		assert_equal(1, cell1.to_i)
		sheet = book.worksheet(2)
    assert_equal('Third Worksheet', sheet.name('latin1'))
    assert_equal(sheet, book.worksheet('Third Worksheet', 'latin1'))
    assert_equal(sheet, book.worksheet("T\0h\0i\0r\0d\0 \0W\0o\0r\0k\0s\0h\0e\0e\0t\0"))
	end
end
