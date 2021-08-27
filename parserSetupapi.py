import os
import datetime
import pymysql as pms

def dt2et(dt) : # 2021/05/29 14:12:32.500
    date_obj = datetime.datetime.strptime(dt, "%Y/%m/%d %H:%M:%S.%f")
    return date_obj.timestamp() *1000 

if __name__ == "__main__" :
    # INSTALL 경로
    installPath = os.getcwd() + r"\c\Windows\appcompat\Programs\Install"
    # Setupapi 경로
    setupapiPath = os.getcwd() + r"\c\Windows\INF"
    # Windows Search/Cortana MRU 경로
    mruPath = os.getcwd() + r"\c\Users\HJun\AppData\Local\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\LocalState\DeviceSearchCache"
    
    

    # 2. setupapi - section
    lfileNames = os.listdir(setupapiPath) 
    
    #    2개의 csv 데이터 생성
    #    section meta data : IDsection, title, start_time, section_end_time, exit_status
    #    section body data : IDsec_body, prefix, event_category, formatted_message, time_stamp, line_num        
    
    secMetas = []
    secBodies = []

    sectionID = 0 
    for fileN in lfileNames : 
        rFlagC = 0
        lFlagC = 0
        logStart = 0
        sectionMeta = []
        sectionBody = []
        bodyID = 1
        with open(setupapiPath+"\\"+fileN,'r', encoding='ANSI') as f :
            lines = f.readlines()
            for line in lines :
                line.strip()
                flag1 = line.split(' ',1)
                if flag1[0] == '>>>' and rFlagC == 0 : # section title 
                    logStart = 1
                    sectionMeta = []
                    
                    rFlagC += 1
                    tmp = flag1[1].split(' ', 1)
                    tmp = tmp[1].strip()
                    tmp = tmp[1:-1]
                    tmp2 = tmp.split('-',1)
                    sectionMeta.append(sectionID)
                    sectionMeta.append(tmp2[0].strip())
                    sectionID += 1 

                elif flag1[0] == '>>>' and rFlagC == 1 : # section start
                    rFlagC = 0
                    tmp = flag1[1].split('Section start', 1)
                    strStime = tmp[1].strip()
                    sEpochTime = dt2et(strStime)
                    sectionMeta.append(int(sEpochTime-32400))    # utc 조정
                elif flag1[0] == '<<<' and lFlagC == 0 : # section end
                    lFlagC = 1
                    tmp = flag1[1].split('Section end',1)
                    strEtime = tmp[1].strip()
                    eEpochTime = dt2et(strEtime)
                    sectionMeta.append(int(eEpochTime-32400))
                elif flag1[0] == '<<<' and lFlagC == 1 : # status
                    logStart = 0
                    lFlagC = 0
                    tmp = flag1[1].split('Exit status:',1)
                    status = tmp[1].strip()
                    sectionMeta.append(status[:-1])
                    secMetas.append(sectionMeta)
                    sectionMeta = []
                elif logStart == 1: 
                    sectionBody = []
                    sectionBody.append(bodyID)            # ID_sec_body = line_num
                    sectionBody.append(flag1[0].strip())  # prefix : none(""), !, !!!
                    flag2 = flag1[1].split(":",1) 
                    sectionBody.append(flag2[0].strip())  # event category : dvi, sig 등
                    sectionBody.append(flag2[1].strip())  # formatted message : dvi, sig 등
                    bodyID += 1
                    secBodies.append(sectionBody)

        
        # csv 콘솔 출력
        print('\n----------------------------------------------------------------------\n')
        print("file name : "+fileN + '\n')
        for metas in secMetas :
            print(metas)
        print(secBodies)
    
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

        for metas in secMetas :
            sql = "INSERT INTO section VALUES %r;" % (tuple(metas),)
            print(sql)
            cursor.execute(sql)
            sqlResult = cursor.fetchall()
            print(sqlResult)

        