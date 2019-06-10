def PushDatabaseinCloudfromXLS(InputPathXLS):
    print('start')
    import pandas as pd 
    DF=pd.read_excel(InputPathXLS)
    import xlrd
    workBook = xlrd.open_workbook(InputPathXLS)
    sheet= workBook.sheet_by_index(0)
    ROWS = sheet.nrows
    COLUMNS =sheet.ncols
    KeysNameforComonInformation=[]
    TotalNumberofSubject=0
    for col in range(1,COLUMNS):
        KeysNameforComonInformation.append(sheet.cell_value(0,col))
    TotalNSL=[]
    for col in range(2,COLUMNS):
        TotalNSL.append(sheet.cell_value(0,col))
    TotalNumberofSubject = int(len(TotalNSL)/5)
    LLforDB=[]
    def FindRowData(i):
        ans=[]
        for col in range(1,COLUMNS):
            ans.append(sheet.cell_value(i,col))
        return ans

    for i in range(1,ROWS):
        LLforDB.append(FindRowData(i))
    def CreateSubjectList(L):
        i=0
        SubjectIDsList=[]
        while(i<len(TotalNSL)):
            SubjectIDsList.append(TotalNSL[i])
            i = i+ TotalNumberofSubject
#         print(SubjectIDsList)
        x=0
        ans=[]
        ans.append(L[0])
        i2=1
        while(i2<len(L)):
            L1=[]
            L1.append(SubjectIDsList[x])
            L1.append(L[i2])
            L1.append(L[i2+1])
            L1.append(L[i2+2])
            L1.append(L[i2+3])
            L1.append(L[i2+4])
            ans.append(L1)
            i2 = i2+5
            x=x+1
        return ans
    LLforMongoMLAB =[]
    for i in LLforDB:
        LLforMongoMLAB.append(CreateSubjectList(i))
    import pymongo
    from pymongo import MongoClient
    Client = pymongo.MongoClient('localhost',27017)
    # Client = pymongo.MongoClient('mongodb://jatinanand345:jatin#123@ds357955.mlab.com:57955/evaluationsystemdb')
    db =  Client.resultconsolidatesystemdb
    students = db.nstudents
    def UpdatedAllInternalValues(L,L1,DB,query):
        ans=[]
        a1=0
        a2=0
        a3=0
        for i in L:
            for j in i:
                if(L1[0].find(j)>-1):
                    for x in i[j]['markswithdate']:
                        for x1 in i[j]['markswithdate'][x]:
                            if(x1.find('internalm')>-1):
                                ans.append([L1[0],L1[1],L1[2],L1[3],L1[4],L1[5],a3,i[j]['markswithdate'][x][x1]])
                                break
                    a1 = a1+ 1
                    break
                a2=a2+1
            a3=a3+1
        return ans
    def UpdateBSONobjectMongoDB(L,DB):
        query={'enrollmentnumber':L[0]}
        StudentPreviousBson = DB.find_one(query)
        StudentPreviousSemesterDict= StudentPreviousBson['semester']
        LObjectSemester= (StudentPreviousSemesterDict.values())
        ans=0
        ANS=[]
        for i in range(1,len(L)):
            ANS.append(UpdatedAllInternalValues(LObjectSemester,L[i],DB,query))
        for j in range(len(ANS)):
            semN = '0'+str(int(ANS[j][0][6])+1)
            PeparCode = ANS[j][0][0]
            ansi1='0'
            ansi2='0'
            ansi3='0'
            ansi4='0'
            ansi5='0'
            if(len(ANS[j][0][1])>0 and len(ANS[j][0][2])>0 and len(ANS[j][0][3])>0 and len(ANS[j][0][4])>0 and len(ANS[j][0][5])>0):
                if(ANS[j][0][1].find('-1')==-1):
                    ansi1=ANS[j][0][1]
                if(ANS[j][0][2].find('-1')==-1):
                    ansi2=ANS[j][0][2]
                if(ANS[j][0][3].find('-1')==-1):
                    ansi3=ANS[j][0][3]
                if(ANS[j][0][4].find('-1')==-1):
                    ansi4=ANS[j][0][4]
                if(ANS[j][0][5].find('-1')==-1):
                    ansi5=ANS[j][0][5]            
            INTERNALUNI = ANS[j][0][-1]
            SUM = str(int(ansi1)+int(ansi2)+int(ansi3)+int(ansi4)+int(ansi5))
            i1='internalm1'
            i2='internalm2'
            i3='internalm3'
            i4='internalm4'
            i5='internalm5'

            KeyString1 = "semester."+semN+"."+PeparCode+"."+i1
            KeyString2 = "semester."+semN+"."+PeparCode+"."+i2
            KeyString3 = "semester."+semN+"."+PeparCode+"."+i3
            KeyString4 = "semester."+semN+"."+PeparCode+"."+i4
            KeyString5 = "semester."+semN+"."+PeparCode+"."+i5
            if(SUM.find(INTERNALUNI)>-1):
                DB.update_one(query,{"$set":{KeyString1:ansi1}})
                DB.update_one(query,{"$set":{KeyString2:ansi2}})
                DB.update_one(query,{"$set":{KeyString3:ansi3}})
                DB.update_one(query,{"$set":{KeyString4:ansi4}})
                DB.update_one(query,{"$set":{KeyString5:ansi5}})
            elif(INTERNALUNI.find('A')>-1 or INTERNALUNI.find('0')>-1):
                DB.update_one(query,{"$set":{KeyString1:'0'}})
                DB.update_one(query,{"$set":{KeyString2:'0'}})
                DB.update_one(query,{"$set":{KeyString3:'0'}})
                DB.update_one(query,{"$set":{KeyString4:'0'}})
                DB.update_one(query,{"$set":{KeyString5:'0'}})

            ans=1
        return ans 

    def FillDatainMongoDBObject(L,DB):
        query={'enrollmentnumber':L[0]}
        f1=0
        NoneType = type(None)
        StudentDict = DB.find_one(query)
        if(type(StudentDict)==NoneType):
    #         if(CreateBSONobjectMongoDB(L,DB)==0):
                f1=1
        else:
            if(UpdateBSONobjectMongoDB(L,DB)==1):
                f1=2
        return f1 
    for i in LLforMongoMLAB:
        FillDatainMongoDBObject(i,students)    
    print('end')
    return 'done'
