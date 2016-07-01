#Scrape Data and export to Excel
#JoeWHoward on GitHub
require 'rubygems'
require 'nokogiri'
require 'writeexcel'

workbook = WriteExcel.new('ruby.xls')

sheet1 = workbook.add_worksheet

sheet1.write(0,0,"A1")
sheet1.write(0,1,"A2")
sheet1.write(1,2,"B3")

temp=0
for i in 2..27
    page = Nokogiri::HTML(open("Page #{i}.html"))
    fullText = page.css("HTML elements")
    num = 0
    counter = 1.0
    
    if i >= 3 
        temp = temp + (fullText.length / 10)
    end
    puts temp
    fullText.pop
    fullText.pop
    fullText.pop
    fullText.each {|x| 
        counter += 0.099999
        if num < 9
            num += 1
        else
            num=0
        end
        if num==0
            sheet1.write(counter.floor + temp,9,x.text)
        elsif num==1
            sheet1.write(counter.floor + temp,0,x.text)
        elsif num==2
            sheet1.write(counter.floor + temp,1,x.text)
        end
    }
end

workbook.close
