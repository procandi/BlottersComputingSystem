# encoding: UTF-8


#匯入模組
if RUBY_VERSION=~/1\.9\../
  require_relative 'ImportSource'
else
  require 'ImportSource'
end


class Main < Sinatra::Base
  #告訴系統在此模組所需要用到的模組
  use ImportSource

  
  #網頁的進入點
  get '/' do   
    #Output Module
    erb :Main
  end
  
  #計算匯圍內的資料
  post '/SearchData' do 
    #setup logger
    @logger=Logger.new("log.txt","daily")
      
    #取得資料庫
    database = SQLite3::Database.new("db/BlottersComputingSystem")
    #執行sql
    searchcode=params[:searchcode]
    begindate=params[:begindate]
    enddate=params[:enddate]
    gatewhere='where 1=1 '
    if searchcode!=nil && searchcode!=''
      gatewhere+="and ItemCode='#{searchcode}' "
    end
    if begindate!=nil && begindate!=''
      gatewhere+="and DT>=#{begindate.gsub(/\//,'')} "
    end
    if enddate!=nil && enddate!=''
      gatewhere+="and DT<=#{enddate.gsub(/\//,'')} "
    end
    sql="select ItemType,KeyNumber,ItemName,Dollar from Blotters #{gatewhere} order by ItemType,KeyNumber,ItemCode,ItemName,Dollar"
    @logger.info(sql)
    rows = database.execute(sql)
    
    #計算並填出資料
    resultdata=''
    accountdata=''
    keynumber=nil
    count=0
    rows.each() do |row|
      if row[3].to_i()>0
        dollar="#{row[3]}, "
      elsif row[3].to_i()<0
        dollar=" ,#{-row[3]}"
      else
        dollar=" , "
      end
      count+=row[3].to_i()
      
      if keynumber==row[0].to_s()
        resultdata+=" ,#{row[1]},#{row[2]},#{dollar}\n"
      else
        keynumber=row[0].to_s()
        accountdata+="#{keynumber},#{count}\n"
        resultdata+="#{keynumber},#{row[1]},#{row[2]},#{dollar}\n"
      end
    end
    resultdata=resultdata[0..resultdata.length-2] if resultdata.length>0
    @result="<table border=1 align='center'><tr><td>科目名稱</td><td>傳票編號</td><td>摘要</td><td>借</td><td>貸</td></tr><tr><td>#{resultdata.gsub(/,/m,'</td><td>').gsub(/\n/m,'</td></tr><tr><td>')}</tr></table>"
    
    #export to excel
    f=File.new("public/Result.xls", "w:utf-8")
    f.write("科目名稱,傳票編號,摘要,借,貸\n")
    f.write(resultdata)
    f.write("\n\n科目名稱,總計\n")
    f.write(accountdata)
    f.close()
    
    database.close if database!=nil
    
    #Output Module
    erb :Main
  end
  
  
  #run sinatra server when this site is unstart
  run! if app_file == $0
end