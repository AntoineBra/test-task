=begin
require "roo"

SUM_LETTER = 'J'
IN_TOTAL_LETTER = 'A'

def count_of_elements(file, first_row)
	cnt = 0
	numArray = Array.new
	while file.cell(IN_TOTAL_LETTER, first_row)!="Итого"
		numArray.push(file.cell(SUM_LETTER, first_row))
		cnt+=1
		first_row+=1
	end 
	cnt_of_each = numArray.each_with_object(Hash.new(0)){ |e, o| o[e] += 1 }
	numArray.clear
	return cnt,cnt_of_each
end



	xlsx = Roo::Excelx.new("../test.xlsx")
	result = count_of_elements(xlsx, 8) 
	count = result[0]
	cnt_of_each_number = result[1]
  	count
	cnt_of_each_number
	cnt_of_each_number.keys
	cnt_of_each_number.values
=end