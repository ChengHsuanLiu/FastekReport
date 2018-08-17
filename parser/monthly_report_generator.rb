##########################
## 每月業績報表
##########################

## 使用方法
## 呼叫程式，後面帶參數譬如 2016.csv, 2017.csv，
## 每個檔案是該年度的年度進銷報表
## ex. ruby parser/monthly_report_generator.rb 2016.csv 2017.csv


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

excelName = Time.now.strftime "每月報表%Y-%m-%d-%H-%M-%S"
book = Spreadsheet::Workbook.new
format = Spreadsheet::Format.new :color => :black,
                                 :size => 16,
                                 :vertical_align => :middle

color_column_fmt = Spreadsheet::Format.new :pattern => 1, :pattern_fg_color => :black, :vertical_align => :middle, :size => 16, :align => :center, :color => :white
color_column_fmt1 = Spreadsheet::Format.new :pattern => 1, :pattern_fg_color => :gray, :vertical_align => :middle, :size => 16, :align => :center, :color => :white

blue_text_fmt = Spreadsheet::Format.new :pattern => 1, :vertical_align => :middle, :size => 16, :align => :center, :color => :blue, :pattern_fg_color => :white

sheet = book.create_worksheet :name => '每月報表'
sheet.default_format = format

sheet[1,0] = "群組欄位"
sheet[1,1] = "索引名稱"
sheet[1,2] = "各分公司"

if( ARGV[0].nil? || ARGV[0].empty? || ARGV.length == 0 )
  return false
end

def default_data
  return {
    "光鈦-桃園" => {
      "1" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "2" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "3" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "4" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "5" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "6" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "7" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "8" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "9" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "10" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "11" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "12" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      }
    },
    "光鈦-台北" => {
      "1" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "2" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "3" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "4" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "5" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "6" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "7" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "8" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "9" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "10" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "11" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "12" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      }
    },
    "光鈦-中壢" => {
      "1" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "2" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "3" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "4" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "5" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "6" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "7" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "8" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "9" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "10" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "11" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "12" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      }
    },
    "光鈦-(竹鈦)" => {
      "1" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "2" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "3" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "4" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "5" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "6" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "7" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "8" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "9" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "10" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "11" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "12" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      }
    },
    "光鈦-台中" => {
      "1" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "2" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "3" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "4" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "5" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "6" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "7" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "8" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "9" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "10" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "11" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "12" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      }
    },
    "光鈦-台南" => {
      "1" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "2" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "3" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "4" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "5" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "6" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "7" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "8" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "9" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "10" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "11" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "12" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      }
    },
    "光鈦高雄" => {
      "1" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "2" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "3" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "4" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "5" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "6" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "7" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "8" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "9" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "10" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "11" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      },
      "12" => {
        "E" => 0,
        "K" => 0,
        "I" => 0,
        "P" => 0,
        "L" => 0
      }
    }
  }
end

result = Hash.new

allMonths = ["1","2","3","4","5","6","7","8","9","10","11","12"]
companies = ["光鈦-台北", "光鈦-中壢", "光鈦-(竹鈦)", "光鈦-台中", "光鈦-台南", "光鈦高雄"]
allCompanies = ["光鈦-桃園", "光鈦-台北", "光鈦-中壢", "光鈦-(竹鈦)", "光鈦-台中", "光鈦-台南", "光鈦高雄", "小計"]
categoryCodes = ["E", "I", "K", "P", "L"]
categoryName = ["電動缸、音圈馬達", "點膠類", "光電類", "馬達HANMARK、TOYO、DELTA", "光寶"]

ARGV.each_with_index do |fileName, file_index|
  result[fileName] = default_data

  totalRevenu = 0

  CSV.foreach("source/#{fileName}", headers: true, encoding: "UTF-8").with_index do |row, fileRow_index|

      sellDate = row[3].strip
      quantity = row[7].strip
      singlePrice = row[9].strip
      singlePriceTotal = row[10].gsub(',', '').to_i
      clientName = row[13].strip
      categoryCode = row[18].strip


      puts "#{sellDate} #{singlePriceTotal} #{clientName} #{categoryCode}"

      begin
        if (categoryCodes.include? categoryCode)

          _month = Date.parse(sellDate).month
          _day = Date.parse(sellDate).day

          if _day > 25
            _month += 1
          end

          if _month == 13
            _month = 1
          end

          _month = _month.to_s

          if (companies.include? clientName)
            originTotal = result[fileName][clientName][_month][categoryCode].to_i
            afterTotal = originTotal + (singlePriceTotal || 0).to_i
            result[fileName][clientName][_month][categoryCode] = afterTotal

            totalRevenu += (singlePriceTotal || 0).to_i
          else
            originTotal = result[fileName]["光鈦-桃園"][_month][categoryCode].to_i
            afterTotal = originTotal + (singlePriceTotal || 0).to_i
            result[fileName]["光鈦-桃園"][_month][categoryCode] = afterTotal

            totalRevenu += (singlePriceTotal || 0).to_i
          end
        end
      rescue Exception => e
        puts e
      end
  end

  puts "totalRevenu: #{totalRevenu}"
  puts "totalRevenu: #{totalRevenu}"
  puts "totalRevenu: #{totalRevenu}"

  allMonths.each_with_index do |month, month_index|
    sheet[0, (file_index*allMonths.count + (month_index + 3))] = "#{fileName.gsub(".csv", "")}年"
    sheet[1, (file_index*allMonths.count + (month_index + 3))] = ( month + " 月" )
  end


  categoryCodes.each_with_index do |cc, index_i|
    allCompanies.each_with_index do |company, index_j|
      allMonths.each_with_index do |month, month_index|

        sheet[( index_i*allCompanies.count + (index_j + 2) ), 0] = cc
        sheet.row(( index_i*allCompanies.count + (index_j + 2) )).set_format(0, blue_text_fmt)
        sheet[( index_i*allCompanies.count + (index_j + 2) ), 1] = categoryName[index_i]
        sheet.row(( index_i*allCompanies.count + (index_j + 2) )).set_format(1, blue_text_fmt)
        sheet[( index_i*allCompanies.count + (index_j + 2) ), 2] = company
        sheet.row(( index_i*allCompanies.count + (index_j + 2) )).set_format(2, blue_text_fmt)

        if company == "小計"

        #   sheet[( index_i*allCompanies.count + (index_j + 1) ),(file_index + 3)] = categorySubtotal

        #   sheet.row(( index_i*allCompanies.count + (index_j + 1) )).set_format(2, color_column_fmt)
        #   sheet.row(( index_i*allCompanies.count + (index_j + 1))).set_format((file_index + 3), color_column_fmt)

        else

          resultTotal = result[fileName][company][month][cc]
          # sheet[( index_i*allCompanies.count + (index_j + 0)),((file_index*(allMonths.length)) + month_index + 3)] = fileName
          # sheet[( index_i*allCompanies.count + (index_j + 1)),((file_index*(allMonths.length)) + month_index + 3)] = month
          sheet[( index_i*allCompanies.count + (index_j + 2)),((file_index*(allMonths.length)) + month_index + 3)] = resultTotal

          # categorySubtotal += resultTotal.to_i

        end


      end
    end
  end
end

[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14].each do |col_index|
  sheet.row(0).set_format(col_index, color_column_fmt)
  sheet.row(1).set_format(col_index, color_column_fmt1)
end

sheet.merge_cells(2, 0, 9, 0)
sheet.merge_cells(10, 0, 17, 0)
sheet.merge_cells(18, 0, 25, 0)
sheet.merge_cells(26, 0, 33, 0)
sheet.merge_cells(34, 0, 41, 0)

sheet.merge_cells(2, 1, 9, 1)
sheet.merge_cells(10, 1, 17, 1)
sheet.merge_cells(18, 1, 25, 1)
sheet.merge_cells(26, 1, 33, 1)
sheet.merge_cells(34, 1, 41, 1)


sheet.rows.map{|row| row.height = 50 if !row.nil?}

for counter in 0..(2 + ARGV.length*12)
  sheet.column(counter).width = 20
end

book.write "results/#{excelName}.xls"
