import os
import datetime
import pymysql as pms


def dt2et(dt) : # 07/09/2021 06:43:43
    date_obj = datetime.datetime.strptime(dt, "%m/%d/%Y %H:%M:%S")
    return int(date_obj.timestamp() *1000 )


if __name__ == "__main__" :
    # INSTALL 경로
    installPath = os.getcwd() + r"\c\Windows\appcompat\Programs\Install"
    # Setupapi 경로
    setupapiPath = os.getcwd() + r"\c\Windows\INF"
    # Windows Search/Cortana MRU 경로
    mruPath = os.getcwd() + r"\c\Users\HJun\AppData\Local\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\LocalState\DeviceSearchCache"
    
    

    # 1. INSTALL 
    lfileNames = os.listdir(installPath) 
    
    #    2개의 csv 데이터 생성
    #    INSTALL meta data : IDInstall, starttime, name, path, binarytype, id, filesize, arpcreate, stoptime
    #    FileCreate data : IDInstall, fcID, path
    
    installFields = ['StartTime','Name','Path','BinaryType', 'Id','FileSize','ArpCreate','StopTime']
    installID = -1
    for fileN in lfileNames : 
        fcID = 0
        idFlag = 0
        acFlag = 0
        installID += 1  
        linstallCsvRow = [str(installID)]
        fcList = []
        with open(installPath+"\\"+fileN,'r', encoding='utf-16-le') as f :
            line = 1
            fcList = []
            lines = f.readlines()
            for line in lines :
                flag1 = line.split('=',1)
                field = flag1[0]  # key
                if field in installFields :
                    if field == "StartTime" or field == "StopTime" : 
                        epochTime = dt2et(flag1[1].strip())
                        linstallCsvRow.append((str)(epochTime))
                        continue
                    # id, arpcreate 같이 여러번 등장하는건 가장 앞의 것으로
                    elif (field == 'Id' and idFlag) or (field == 'ArpCreate' and acFlag):
                        continue
                    elif field == 'BinaryType' :
                        linstallCsvRow.append(str(flag1[1].strip()))
                        continue 
                    linstallCsvRow.append(flag1[1].strip())
                # FileCreate의 경우 csv의 list로 
                elif field == 'FileCreate' :
                    fcList.append(",".join([str(installID),str(fcID),flag1[1].strip()]))
                    fcID +=1
            
            strInstallCsvRow = ",".join(linstallCsvRow)
            
        # csv 콘솔 출력
        print('\n----------------------------------------------------------------------\n')
        print("file name : "+fileN + '\n')
        print(strInstallCsvRow)
        print(fcList)

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

    sql = "INSERT INTO install VALUES %r;" % (tuple(linstallCsvRow),)
    print(sql)
    cursor.execute(sql)
    sqlResult = cursor.fetchall()
    print(sqlResult)


    