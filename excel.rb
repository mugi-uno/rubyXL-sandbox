require 'pry'
require 'chroma'
require 'rubyXL'
require 'rubyXL/convenience_methods'

workbook = RubyXL::Parser.parse(Pathname.new('./template.xlsx'))

sheet = workbook.worksheets[0]

puts "num: #{sheet[3][2].value}"
puts "str: #{sheet[4][2].value}"
puts "dat: #{sheet[5][2].value}"

data = [
  { a: 11, b: 12, c: 13, d: 14 },
  { a: 21, b: 22, c: 23, d: 24 },
  { a: 31, b: 32, c: 33, d: 34 },
  { a: 41, b: 42, c: 43, d: 44 },
  { a: 51, b: 52, c: 53, d: 54 },
]

# 行挿入
row_num = 8

(data.size - 1).times do
  sheet.insert_row(row_num + 1)
end

def copy_cell(from, to)
  to.change_fill(from.fill_color.slice(0, 6))
  to.change_font_name(from.font_name)
  to.change_font_size(from.font_size)
  to.change_font_color(from.font_color.paint.to_hex.slice(1, 7))
  to.change_font_italics(from.is_italicized == true)
  to.change_font_bold(from.is_bolded == true)
  to.change_font_underline(from.is_underlined == true)
  to.change_font_strikethrough(from.is_struckthrough == true)
end

(data.size - 1).downto(0).each do |row_index|
  (0..65_535).each do |num|
    break unless sheet[row_num][num]
    break if row_index == 0

    sheet.insert_cell(row_num + row_index, num, sheet[row_num][num].value)

    copy_cell(sheet[row_num][num], sheet[row_num + row_index][num])
  end
end

# 書き込み
workbook.write(Pathname.new('./new.xlsx'))
