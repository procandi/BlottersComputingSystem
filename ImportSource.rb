# encoding: UTF-8


class ImportSource < Sinatra::Base
  #Import Source Excel
  get '/ImportSource' do
    @uploadok=0
    
    #Output Module
    erb :ImportSource
  end
  
  #Upload Source Excel
  post "/ImportSource" do
    begin
      #setup logger
      @logger=Logger.new("log.txt","daily")
      
      #Upload File
      tempfile = params['file'][:tempfile]
      filename = params['file'][:filename]
        
      if filename[-4..-1]=='.xls'
        FileUtils.cp(tempfile.path, "public/#{filename}")
        @uploadok=1
      else
        @uploadok=-1
      end
      
      #Read Excel, and Write to DB
      if @uploadok==1
        #Backup DB
        FileUtils.cp("db/BlottersComputingSystem","db/BlottersComputingSystem_#{DateTime.now.strftime('%Y%m%d%H%M%S')}")
        
        #Init DB variant
        database = SQLite3::Database.new("db/BlottersComputingSystem")
        
        #Read Excel
        flag=false     
        @workbook = Spreadsheet.open("public/#{filename}")
        @workbook.worksheets.each() do |worksheet|
          #@worksheet = @workbook.worksheet(0)
          0.upto worksheet.last_row_index do |index|
            row = worksheet.row(index)
            
            if row[3]!=nil && (row[3]=='借' || row[3]=='貸') && !flag #用以判斷此項是否為資料，還是其它的標題
              #p row[0]  #KeyNumber,DTD,DTM,DTY
              #p row[1]  #ItemCode
              #p row[2]  #ItemType
              #p row[3]  #借或貸
              #p row[4]  #Add Dollar
              #p row[5]  #Cut Dollar            
             
              unless row[0]==nil
                @keynumber=row[0].to_s()
                @dtd=@keynumber[5..6].to_i()
                @dtm=@keynumber[3..4].to_i()
                @dty=@keynumber[0..2].to_i()+1911
                @dtd=@dty.to_s()+@dtm.to_s()+@dtd.to_s()
              end
              @itemcode=row[1].to_s()
              @itemtype=row[2].to_s().gsub(/\'/m,'\'\'')
              flag=true
              if row[3].to_s()=='借'
                unless row[4]==nil
                  @dollar=row[4].to_i()
                else
                  @dollar=0
                end
              else
                unless row[5]==nil
                  @dollar=-row[5].to_i()
                else
                  @dollar=0
                end
              end
            elsif flag
              #p row[2]  #ItemName
              @itemname=row[2].to_s().gsub(/\'/m,'\'\'')
              
              #Write to DB
              sql="insert into Blotters(DT,DTD,DTM,DTY,KeyNumber,ItemType,ItemCode,ItemName,Dollar) values(#{@dt},#{@dtd},#{@dtm},#{@dty},'#{@keynumber}','#{@itemtype}','#{@itemcode}','#{@itemname}',#{@dollar})"
              @logger.info(sql)
              database.execute(sql)
          
              flag=false
            else
              flag=false
            end
          end
        end
      end
      
      database.close if database!=nil
      
      #Output Module
      erb :ImportSource
    rescue => e
      @uploadok=1
      database.close if database!=nil
      
      '上傳發生錯誤'
    end
  end
end