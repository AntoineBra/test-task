require "roo"
source = 'http://some.remote.host/barchart.xml'


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


SCHEDULER.every '10s', :first_in => 0 do |job|
	xlsx = Roo::Excelx.new("./test.xlsx")
	result = count_of_elements(xlsx, 8)
	total_cnt_of_elements = result[0]
	cnt_of_each_hash = result[1]
	sorted_cnt_of_each_hash = cnt_of_each_hash.sort.to_h

  data = [
      label: 'Dataset',
      data: Array.new(sorted_cnt_of_each_hash.length) { |i| sorted_cnt_of_each_hash.values[i] },
      backgroundColor: [ 'rgba(255, 99, 132, 0.2)' ] * sorted_cnt_of_each_hash.keys.length,
      borderColor: [ 'rgba(255, 99, 132, 1)' ] * sorted_cnt_of_each_hash.keys.length,
      borderWidth: 1,
  ]

  send_event('cnt_wid', { title: "Elements: ", text: total_cnt_of_elements })
  send_event('each_cnt_wid', { title: "Count of each: ", items: cnt_of_each_hash.sort.map{|k,v| {label: k, value: v}} })
  send_event('barchart', { labels: sorted_cnt_of_each_hash.keys, datasets: data })
end
