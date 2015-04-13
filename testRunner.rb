require "selenium-webdriver"
require "simple-spreadsheet"
require "colorize"
require "csv"

config = Hash.new
arry = Array.new

$sheet = SimpleSpreadsheet::Workbook.read("testCase.xlsx")
$URL = $sheet.cell(1, 2) 

lastRow = $sheet.last_row
startRow = 3

def read_csv(fileName)
  csvData = CSV.read(fileName, headers:true)
  return csvData
end

def execute_steps(rowNum, stop, row="")
  $driver = Selenium::WebDriver.for :firefox
  $driver.manage.window.maximize
  $driver.navigate.to $URL
  $testResultFlag = false
  
  begin
    while (rowNum < stop) do
      rowNum += 1
    
      ######### PERFORM ACTIONS #########
    
      if($sheet.cell(rowNum, 1) == "Type")
        if($sheet.cell(rowNum, 2) == "id")
           element = $driver.find_element(:id, $sheet.cell(rowNum, 3))
         elsif($sheet.cell(rowNum, 2) == "xpath")
           element = $driver.find_element(:xpath, $sheet.cell(rowNum, 3))
         end
       
         if($sheet.cell(rowNum, 4).start_with?("#"))
           key = $sheet.cell(rowNum, 4).split("#").last
           element.send_keys row[key]
         else
           element.send_keys $sheet.cell(rowNum, 4)
         end
     
       elsif($sheet.cell(rowNum, 1) == "Click")
         if($sheet.cell(rowNum, 2) == "id")
           element = $driver.find_element(:id, $sheet.cell(rowNum, 3))
         elsif($sheet.cell(rowNum, 2) == "xpath")
           element = $driver.find_element(:xpath, $sheet.cell(rowNum, 3))
         end
         element.click
     
       elsif($sheet.cell(rowNum, 1) == "Select")
         if($sheet.cell(rowNum, 2) == "id")
           element = $driver.find_element(:id, $sheet.cell(rowNum, 3))
         elsif($sheet.cell(startRow, 2) == "xpath")
           element = $driver.find_element(:xpath, $sheet.cell(rowNum, 3))
         end
         select_list = Selenium::WebDriver::Support::Select.new(element)
       
         if($sheet.cell(rowNum, 4).start_with?("#"))
           key = $sheet.cell(rowNum, 4).split("#").last
           select_list.select_by(:text, row[key])
         else
           select_list.select_by(:text, $sheet.cell(rowNum, 4))
         end

       elsif($sheet.cell(rowNum, 1) == "Submit")
         element.submit
       end
     
       ######### VALIDATE RESULTS #########
     
       if($sheet.cell(rowNum, 5) != nil)
         if($sheet.cell(rowNum, 5) == "id")
           element = $driver.find_element(:id, $sheet.cell(rowNum, 6))
         elsif($sheet.cell(rowNum, 5) == "xpath")
           element = $driver.find_element(:xpath, $sheet.cell(rowNum, 6))
         end
       
         if($sheet.cell(rowNum, 7) == "equal")
           if(element.text == $sheet.cell(rowNum, 8))
             $testResultFlag = true
           end
         
         elsif($sheet.cell(rowNum, 7) == "contains")
           if(element.text.include? $sheet.cell(rowNum, 8))
             $testResultFlag = true
           end

         elsif($sheet.cell(rowNum, 7) == nil)
           if($driver.find_elements(:id, $sheet.cell(rowNum, 6)).size()) > 0
             $testResultFlag = true
           end
         end
       end
    end
  rescue
    $testResultFlag = false
  end
  $driver.quit
end

def execute_test(rowNum, stop)
  testName = $sheet.cell(rowNum, 1)
  
  csvData = ""
  if($sheet.cell(rowNum, 2) != nil)
    csvData = read_csv($sheet.cell(rowNum, 2))
    csvData.each do |row|
      puts "Test \"" + testName + "\" has been started with data:"
      puts row.to_s
      execute_steps(rowNum, stop, row)
      
      if ($testResultFlag)
        puts "PASSED".on_green
        puts ""
      else
        puts "FAILED".on_red
        puts ""
      end
    end
  else
    puts "Test \"" + testName + "\" has been started"
    execute_steps(rowNum, stop)
    
    if ($testResultFlag)
      puts "PASSED".on_green
      puts ""
    else
      puts "FAILED".on_red
      puts ""
    end
  end
end


while startRow <= lastRow do
  min = startRow
  while($sheet.cell(startRow, 1) != nil) do
    startRow += 1
  end
  max = startRow
  
  execute_test(min, max)
  startRow += 1
end