import os
import json
import pymysql as pms
import datetime

EPOCH_AS_FILETIME = 116444736000000000  # January 1, 1970 as filetime
HUNDREDS_OF_NS = 10000000

def to_datetime(filetime: int) -> datetime:
	"""
	Converts a Windows filetime number to a Python datetime. The new
	datetime object is timezone-naive but is equivalent to tzinfo=utc.
	"""

	# Get seconds and remainder in terms of Unix epoch
	s, ns100 = divmod(filetime - EPOCH_AS_FILETIME, HUNDREDS_OF_NS)
	# Convert to datetime object, with remainder as microseconds.
	return datetime.datetime.utcfromtimestamp(s).replace(microsecond=(ns100 // 10))

if __name__ == "__main__" :
    # INSTALL 경로
    installPath = os.getcwd() + r"\c\Windows\appcompat\Programs\Install"
    # Setupapi 경로
    setupapiPath = os.getcwd() + r"\c\Windows\INF"
    # Windows Search/Cortana MRU 경로
    mruPath = os.getcwd() + r"\c\Users\HJun\AppData\Local\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\LocalState\DeviceSearchCache"
    
    
   
    
    # 3. MRU
    lfileNames = os.listdir(mruPath) 
    
    #    2개의 csv 데이터 생성
    #    MRU data : IDsection, title, start_time, section_end_time, exit_status
    #    jump list data : IDsec_body, prefix, event_category, formatted_message, time_stamp, line_num  
    id = -1
    jumplists = []
    mruRows = []
    for fileN in lfileNames : 
        with open(mruPath+"\\"+fileN,'r', encoding='utf8') as f :
            MRUjson = json.load(f)
            for jsons in MRUjson : 
                id += 1
                jlID = 0
                mruRowList = [id]
                jlRowList = [id]
                fileTime = fileN.replace("AppCache", '').replace(".txt",'')
                
                convDataTime = to_datetime(int(fileTime))
                convEpochTime = int(convDataTime.timestamp())

                mruRowList.append(jsons["System.FileExtension"]["Value"])
                mruRowList.append(jsons["System.Software.ProductVersion"]["Value"])
                mruRowList.append(jsons["System.Kind"]["Value"])
                mruRowList.append(jsons["System.ParsingName"]["Value"])
                mruRowList.append(jsons["System.Software.TimesUsed"]["Value"])
                mruRowList.append(jsons["System.Tile.Background"]["Value"])
                mruRowList.append(jsons["System.AppUserModel.PackageFullName"]["Value"])
                mruRowList.append(jsons["System.Identity"]["Value"])
                mruRowList.append(jsons["System.FileName"]["Value"])
                mruRowList.append(jsons["System.ConnectedSearch.VoiceCommandExamples"]["Value"])
                mruRowList.append(jsons["System.ItemType"]["Value"])
                mruRowList.append(jsons["System.DateAccessed"]["Value"])
                mruRowList.append(jsons["System.Tile.EncodedTargetPath"]["Value"])
                mruRowList.append(jsons["System.Tile.SmallLogoPath"]["Value"])
                mruRowList.append(jsons["System.ItemNameDisplay"]["Value"])
                mruRowList.append(convEpochTime) # epoch time
                mruRows.append(mruRowList)
                jl = json.loads(jsons["System.ConnectedSearch.JumpList"]["Value"]) # jumplist
                for js in jl : 
                    for j in js["Items"] :
                        jlRowList = [id] 
                        jlRowList.append(jlID)
                        jlID += 1
                        try : 
                            jlRowList.append(j["Type"])
                            jlRowList.append(j["Name"])
                            jlRowList.append(j["Path"])
                            jlRowList.append(j["Date"])
                            jlRowList.append(j["Points"])
                        except : 
                            jlRowList.append("")
                        jumplists.append(jlRowList)
    

             
    # csv 콘솔 출력
        print('\n----------------------------------------------------------------------\n')
        print("file name : "+fileN + '\n')
        # for row in mruRows :
            # print(row)
        # print(jumplists)
    

# DB에 넣기
    # 
    mydb = pms.connect(
                    user='root', 
                    passwd='0000', 
                    host='127.0.0.1', 
                    db='mydb', 
                    charset='utf8'
                    )
    cursor = mydb.cursor(pms.cursors.DictCursor)
    for row in mruRows :
        print(len(row))
        sql = "INSERT INTO cortanamru VALUES %r;" % (tuple(row),)
        print(sql)
        cursor.execute(sql)
        sqlResult = cursor.fetchall()
        print(sqlResult)
            