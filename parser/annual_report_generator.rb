##########################
## 年度業績報表
##########################

## 使用方法
## 呼叫程式，後面帶參數譬如 2016.csv, 2017.csv，
## 每個檔案是該年度的年度進銷報表
## ex. ruby parser/annual_report_generator.rb 2016.csv 2017.csv


## 產品分類
# E 電動缸 / 音圈馬達
# K 光電類
# I 點膠類
# P 馬達 Hanmark / TOYO / DELTA
# L 光寶

# 分公司
# 桃園總公司
# 台北：光鈦-台北
# 中壢：光鈦-中壢
# 新竹：光鈦-(竹鈦)
# 台中：光鈦-台中
# 台南：光鈦-台南
# 高雄：光鈦高雄

require 'csv'
require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

excelName = Time.now.strftime "年報表%Y-%m-%d-%H-%M-%S"
book = Spreadsheet::Workbook.new
format = Spreadsheet::Format.new :color => :black,
                                 :size => 16,
                                 :vertical_align => :middle

color_column_fmt = Spreadsheet::Format.new :pattern => 1, :pattern_fg_color => :gray, :vertical_align => :middle, :size => 16, :align => :center, :color => :white

sheet = book.create_worksheet :name => '年報表'
sheet.default_format = format

sheet[0,0] = "群組欄位"
sheet[0,1] = "索引名稱"
sheet[0,2] = "各分公司"

sheet.row(0).set_format(0, color_column_fmt)
sheet.row(0).set_format(1, color_column_fmt)
sheet.row(0).set_format(2, color_column_fmt)

if( ARGV[0].nil? || ARGV[0].empty? || ARGV.length == 0 )
  return false
end

def default_data
  return {
    "光鈦-桃園" => {
      "E" => 0,
      "K" => 0,
      "I" => 0,
      "P" => 0,
      "L" => 0
    },
    "光鈦-台北" => {
      "E" => 0,
      "K" => 0,
      "I" => 0,
      "P" => 0,
      "L" => 0
    },
    "光鈦-中壢" => {
      "E" => 0,
      "K" => 0,
      "I" => 0,
      "P" => 0,
      "L" => 0
    },
    "光鈦-(竹鈦)" => {
      "E" => 0,
      "K" => 0,
      "I" => 0,
      "P" => 0,
      "L" => 0
    },
    "光鈦-台中" => {
      "E" => 0,
      "K" => 0,
      "I" => 0,
      "P" => 0,
      "L" => 0
    },
    "光鈦-台南" => {
      "E" => 0,
      "K" => 0,
      "I" => 0,
      "P" => 0,
      "L" => 0
    },
    "光鈦高雄" => {
      "E" => 0,
      "K" => 0,
      "I" => 0,
      "P" => 0,
      "L" => 0
    }
  }
end

result = Hash.new
companies = ["光鈦-台北", "光鈦-中壢", "光鈦-(竹鈦)", "光鈦-台中", "光鈦-台南", "光鈦高雄"]
allCompanies = ["光鈦-桃園", "光鈦-台北", "光鈦-中壢", "光鈦-(竹鈦)", "光鈦-台中", "光鈦-台南", "光鈦高雄", "小計"]
categoryCodes = ["E", "I", "K", "P", "L"]
categoryName = ["電動缸、音圈馬達", "點膠類", "光電類", "馬達HANMARK、TOYO、DELTA", "光寶"]



ARGV.each_with_index do |fileName, index_1|
  result[fileName] = default_data

  sheet[0,(index_1 + 3)] = fileName.gsub('.csv', '')
  sheet.row(0).set_format((index_1 + 3), color_column_fmt)

  annualTotal = 0
  # fileName ex. 2016.csv
  CSV.foreach("source/#{fileName}", headers: true, encoding: "UTF-8").with_index do |row, index_2|
    # if index_2 < 3

      sellDate = row[3].strip
      quantity = row[7].strip
      singlePrice = row[9].strip
      singlePriceTotal = row[10].to_s.gsub(',', '').to_i
      clientName = row[13].strip
      categoryCode = row[18].strip
      puts "#{sellDate} #{singlePriceTotal} #{clientName} #{categoryCode}"

      begin
        if (categoryCodes.include? categoryCode)
          if (companies.include? clientName)
            originTotal = result[fileName][clientName][categoryCode].to_i
            afterTotal = originTotal + (singlePriceTotal || 0).to_i
            result[fileName][clientName][categoryCode] = afterTotal
          else
            originTotal = result[fileName]["光鈦-桃園"][categoryCode].to_i
            afterTotal = originTotal + (singlePriceTotal || 0).to_i
            result[fileName]["光鈦-桃園"][categoryCode] = afterTotal
          end
        end
      rescue Exception => e
        puts e
      end
    # end
  end

  puts result
  categoryCodes.each_with_index do |cc, index_i|
    categorySubtotal = 0

    allCompanies.each_with_index do |company, index_j|
      sheet[( index_i*allCompanies.count + (index_j + 1) ), 0] = cc
      sheet[( index_i*allCompanies.count + (index_j + 1) ), 1] = categoryName[index_i]
      sheet[( index_i*allCompanies.count + (index_j + 1) ), 2] = company

      if company == "小計"
        sheet[( index_i*allCompanies.count + (index_j + 1) ),(index_1 + 3)] = categorySubtotal

        sheet.row(( index_i*allCompanies.count + (index_j + 1) )).set_format(2, color_column_fmt)
        sheet.row(( index_i*allCompanies.count + (index_j + 1))).set_format((index_1 + 3), color_column_fmt)
      else
        resultTotal = result[fileName][company][cc]
        sheet[( index_i*allCompanies.count + (index_j + 1)),(index_1 + 3)] = resultTotal

        categorySubtotal += resultTotal.to_i
      end
    end

    annualTotal += categorySubtotal
  end

  sheet[(allCompanies.count*categoryCodes.count + 1), 2] = "總計"
  sheet[(allCompanies.count*categoryCodes.count + 1), (index_1 + 3)] = annualTotal
end

# sheet.merge_cells(start_row, start_col, end_row, end_col)
sheet.merge_cells(1, 0, 8, 0)
sheet.merge_cells(9, 0, 16, 0)
sheet.merge_cells(17, 0, 24, 0)
sheet.merge_cells(25, 0, 32, 0)
sheet.merge_cells(33, 0, 40, 0)

sheet.merge_cells(1, 1, 8, 1)
sheet.merge_cells(9, 1, 16, 1)
sheet.merge_cells(17, 1, 24, 1)
sheet.merge_cells(25, 1, 32, 1)
sheet.merge_cells(33, 1, 40, 1)


sheet.rows.map{|row| row.height = 50 if !row.nil?}

for counter in 0..(2 + ARGV.length)
  sheet.column(counter).width = 30
end


book.write "results/#{excelName}.xls"
