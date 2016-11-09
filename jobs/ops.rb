# :first_in sets how long it takes before the job is first run. In this case, it is run immediately
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

data = Array.new

SCHEDULER.every '60s', :first_in => 0 do |job|
	xlsx = Roo::Excelx.new("./test.xlsx")
	result = count_of_elements(xlsx, 8) 
  	send_event('my_widget', { val: result[0] })
  	send_event('my_widget2', {items: result[1].sort.map{|k,v| {label: k, value: v}} })
  	
  	data.clear
  	result[1].sort.each do |digit, count|
  		data << { "x": digit, "y": count }
	end
	data = data[0..-2]
  send_event('my_wid', { points: data })

end