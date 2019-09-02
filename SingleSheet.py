import xlsxwriter
import os,string
import time
import os.path
import win32file
import win32con
import datetime
import xlrd
from datetime import date
import matplotlib.pyplot as plt
import operator

objfileList = xlsxwriter.Workbook("FilesList.xlsx")
fileList=objfileList.add_worksheet()
anaobjfileList =xlsxwriter.Workbook("AnalysisFilesList.xlsx")
anafileList=anaobjfileList.add_worksheet()
folderobjfileList =xlsxwriter.Workbook("Folderwise.xlsx")
folderfileList=folderobjfileList.add_worksheet()
notUsedfileList =xlsxwriter.Workbook("NotUsedFile.xlsx")
notfileList=notUsedfileList.add_worksheet()
filescount=0
def monthCheck(alaMonth):
            if(alaMonth=='Jan'):
                month=1
            elif(alaMonth=='Feb'):
                month=2
            elif(alaMonth=='Mar'):
                month=3
            elif(alaMonth=='Apr'):
                month=4
            elif(alaMonth=='May'):
                month=5
            elif(alaMonth=='Jun'):
                month=6
            elif(alaMonth=='Jul'):
                month=7
            elif(alaMonth=='Aug'):
                month=8
            elif(alaMonth=='Sep'):
                month=9
            elif(alaMonth=='Oct'):
                month=10
            elif(alaMonth=='Nov'):
                month=11
            elif(alaMonth=='Dec'):
                month=12
            return month
class Analysis:
    def duplicateFile(self,total_count):
        print("DuplicateFile process started")
        count_display=[]
        start=time.time()
        total_count_update=total_count+1
        location = (r"FilesList.xlsx")
        file = xlrd.open_workbook(location) 
        sheet = file.sheet_by_index(0) 
        sheet.cell_value(0, 0)
        size=0
        count=1
        count2=1
        objdupfileList = xlsxwriter.Workbook("Duplicatefiles.xlsx")
        dupfileList=objdupfileList.add_worksheet('Sheet 1')
        dupfileList2=objdupfileList.add_worksheet('Sheet 2')
        dupfileList.write(0,0,'S.No')
        dupfileList.write(0,1,'File Name')
        dupfileList.write(0,2,'Location')
        dupfileList2.write(0,0,'S.No')
        dupfileList2.write(0,1,'File Name')
        dupfileList2.write(0,2,'Location')
        dupfileList.set_column(1, 1, 30)
        dupfileList.set_column(2, 2, 170)
        dupfileList2.set_column(1, 1, 30)
        dupfileList2.set_column(2, 2, 170)
        row=1
        row2=1
        for i in range(1,total_count):
            
            iextension=sheet.cell_value(i,2)
            icreated=sheet.cell_value(i,5)
            imodified=sheet.cell_value(i,4)
            iaccess=sheet.cell_value(i,3)
            ifilename=sheet.cell_value(i,1)
            isize=sheet.cell_value(i,6)
            ihidden=sheet.cell_value(i,8)
        
            flag=False
            flag2=False
            for j in range(i+1,total_count_update):
                jextension=sheet.cell_value(j,2)
                jcreated=sheet.cell_value(j,5)
                jfilename=sheet.cell_value(j,1)
                jsize=sheet.cell_value(j,6)
                jmodified=sheet.cell_value(j,4)
                jaccess=sheet.cell_value(j,3)
                jhidden=sheet.cell_value(j,8)
                if((iextension==jextension) and (ifilename==jfilename) and (isize==jsize)and(imodified==jmodified)and(ihidden=='No')and(jhidden=='No')):
                    if(flag==False):
                        dupfileList.write(row, 0, count)
                        count+=1
                        flag=True
                        
                    dupfileList.write(row, 1, ifilename)
                    dupfileList.write(row, 2, sheet.cell_value(i,12))
                    dupfileList.write(row+1, 1, jfilename)
                    dupfileList.write(row+1, 2, sheet.cell_value(j,12))
                    location=sheet.cell_value(j,12)
                    count_display.append(location[0])
                    locationi=sheet.cell_value(i,12)
                    count_display.append(locationi[0])
                    size+=jsize
                    row=row+2
                    
        

        
        count=0
        for i in range(1,total_count):
            size=sheet.cell_value(i,6)
            if(size==0):
                count+=1
                dupfileList2.write(count,0,count)
                dupfileList2.write(count,1,sheet.cell_value(i,1))
                dupfileList2.write(count,2,sheet.cell_value(i,12))
            
                    
                
                
                    
                
        print("\n\nBecause of duplicate files you wasted",size," bytes")    
        end=time.time()
        print("DuplicateFile process completed")
        print("DuplicateFile process time taken is",end-start)
        objdupfileList.close()
        #to display excel sheet
        file = r"Duplicatefiles.xlsx"
        os.startfile(file)

        dic_display={}
        for i in range(0,len(count_display)):
            if count_display[i] not in dic_display:
                dic_display[count_display[i]]=count_display.count(count_display[i])

        x=[]
        count=[]
        for key,value in dic_display.items():
            x.append(key)
            count.append(value)

        plt.pie(count, labels=x,autopct='%1.1f%%',radius=5 ,shadow=True, startangle=140)
        plt.title("Percentage of duplicate files in different drives")
        plt.axis('equal')
        plt.show()

    def year_month_percentage(self,total_count,year):
        print("Month wise analysis process started")
        start=time.time()
        location = (r"FilesList.xlsx") 
        file = xlrd.open_workbook(location) 
        sheet = file.sheet_by_index(0) 
        sheet.cell_value(0, 0)
        pert_month={}
        for i in range(0,len(year)):
            print(year[i])
        need_year=int(input('Enter the year for which you need month wise analysis: '))
        if need_year in year:
            total_count=total_count+1
            for i in range(1,total_count):
                CreatedDate=sheet.cell_value(i,5)
                creDate,creMonth,creDay,creTime,creYear=CreatedDate.split()
                creYear=int(creYear)
                if(creYear==need_year):
                    if(creMonth not in pert_month):
                        pert_month.update({creMonth:0})
                        count=0
                        for j in range(1,total_count):
                            CreatedDatej=sheet.cell_value(j,5)
                            creDatej,creMonthj,creDayj,creTimej,creYearj=CreatedDatej.split()
                            creYearj=int(creYearj)
                            if(need_year==creYearj and creMonthj==creMonth):
                                count=count+1
                        pert_month[creMonth]=count
            end=time.time()
            print("Month wise analysis process completed")
            print("Month wise analysis process time taken is",end-start)
            x=[]
            count_month=[]
            for key,value in pert_month.items():
                x.append(key)
                count_month.append(value)

        
            plt.pie(count_month, labels=x,autopct='%1.1f%%',radius=5 ,shadow=True, startangle=140)
            need_year=str(need_year)
            plt.title("Month wise analysis for the year "+need_year)
            plt.axis('equal')
            plt.show()               
            return 0    
            
        else:
            print('You have entered the wrong year')
            return 1
    def year_percentage(self,total_count):
        print("Year wise analysis process started")
        start=time.time()
        location = (r"FilesList.xlsx") 
        file = xlrd.open_workbook(location) 
        sheet = file.sheet_by_index(0) 
        sheet.cell_value(0, 0)
        pert_year={}
        total_count=total_count+1
        for i in range(1,total_count):
            CreatedDate=sheet.cell_value(i,5)
            creDate,creMonth,creDay,creTime,creYear=CreatedDate.split()
            creYear=int(creYear)
            if creYear not in pert_year:
                pert_year.update({creYear:0})
                count=0
                for j in range(1,total_count):
                    CreatedDatej=sheet.cell_value(j,5)
                    creDatej,creMonthj,creDayj,creTimej,creYearj=CreatedDatej.split()
                    creYearj=int(creYearj)
                    if(creYear==creYearj):
                        count=count+1
                pert_year[creYear]=count
        end=time.time()
        print("Year wise analysis process completed")
        print("Year wise analysis process time taken is",end-start)
        x=[]
        count_year=[]
        for key,value in pert_year.items():
            x.append(key)
            count_year.append(value)

        
        plt.pie(count_year, labels=x,autopct='%1.1f%%',radius=5 ,shadow=True, startangle=140)
        plt.title("Year wise analysis")
        plt.axis('equal')
        plt.show()
        return x
                    

            
        
    def percentage(self,total_count):
        print("Percentage process started")
        start=time.time()
        location = (r"FilesList.xlsx") 
        file = xlrd.open_workbook(location) 
        sheet = file.sheet_by_index(0) 
        sheet.cell_value(0, 0)
        pert={}
        pert_count={}
        total_count=total_count+1
        for i in range(1,total_count): 
                extension=sheet.cell_value(i,2)
                if extension not in pert:
                    pert.update({extension:0})
                    count=0
                    for j in range(1,total_count):
                        get_extension=sheet.cell_value(j,2)
                        if(get_extension==extension):
                            count=count+1
                    pert[extension]=count
        end=time.time()
        print("Percentage process completed")
        print("Percentage process time taken is",end-start)
        pert_sorted = sorted(pert, key=pert.get, reverse=True)
        for r in pert_sorted:
            pert_count.update({r:pert[r]})

        
        x=[]
        count=[]
        i=0
        for key,value in pert_count.items():
            x.append(key)
            count.append(value)
            i=i+1
            if(i==15):
                break
        x_pos = [i for i, _ in enumerate(x)]
        plt.bar(x_pos, count, edgecolor='blue')
        plt.xlabel("Extension")
        plt.ylabel("Count")
        plt.title("Top 15 files count based on extension")
        plt.xticks(x_pos, x)
        plt.minorticks_on()
        plt.grid(which='major', linestyle='-', linewidth='0.5', color='red')
        plt.grid(which='minor', linestyle=':', linewidth='0.5', color='black')
        plt.show()


        
 
        plt.pie(count, labels=x,autopct='%1.1f%%',radius=5 ,shadow=True, startangle=140)
        plt.title("Top 15 files percertange count based on extension")
        plt.axis('equal')
        plt.show()

    def DateDifference(self,count):
        anafileList.write(0,0,'S.No')
        anafileList.write(0,1,'Access Date Difference')
        anafileList.write(0,2,'Modified Date Difference')
        anafileList.write(0,3,'Created Date Difference')
        anafileList.set_column(0, 3, 30)
        print("Date Difference process started")
        start=time.time()
        location=(r"FilesList.xlsx")
        file=xlrd.open_workbook(location)
        sheet=file.sheet_by_index(0)
        sheet.cell_value(0,0)
        now=datetime.datetime.now()
        count=count+1
        for i in range(1,count):
            AccessDate=sheet.cell_value(i,3)
            accDate,accMonth,accDay,accTime,accYear=AccessDate.split()
            accMonth=monthCheck(accMonth)

            ModifiedDate=sheet.cell_value(i,4)
            modDate,modMonth,modDay,modTime,modYear=ModifiedDate.split()
            modMonth=monthCheck(modMonth)

            CreatedDate=sheet.cell_value(i,5)
            creDate,creMonth,creDay,creTime,creYear=CreatedDate.split()
            creMonth=monthCheck(creMonth)

            
            accYear=int(accYear)
            accDay=int(accDay)

            modYear=int(modYear)
            modDay=int(modDay)

            creYear=int(creYear)
            creDay=int(creDay)
            
            
            acc_firstDate = datetime.date(accYear,accMonth,accDay)
            acc_lastDate = datetime.date(now.year,now.month,now.day)
            acc_difference=acc_lastDate-acc_firstDate
            anafileList.write(i,0,i)
            anafileList.write(i,1,acc_difference.days)


            mod_firstDate = datetime.date(modYear,modMonth,modDay)
            mod_lastDate = datetime.date(now.year,now.month,now.day)
            mod_difference=mod_lastDate-mod_firstDate
            anafileList.write(i,2,mod_difference.days)

            cre_firstDate = datetime.date(creYear,creMonth,creDay)
            cre_lastDate = datetime.date(now.year,now.month,now.day)
            cre_difference=cre_lastDate-cre_firstDate
            anafileList.write(i,3,cre_difference.days)

        end=time.time()
        print("Date Difference process completed")
        print("Date Difference process time taken is",end-start)

    def notUsedFile(self,count):
        count_display=[]
        sample_extension=['mp3','mp4','mkv','flv','gif','gifv','mp2']
        count=count+1
        location = (r"AnalysisFilesList.xlsx") 
        file = xlrd.open_workbook(location) 
        sheet = file.sheet_by_index(0) 
        sheet.cell_value(0, 0)

        location1 = (r"FilesList.xlsx") 
        file1 = xlrd.open_workbook(location1) 
        sheet1 = file1.sheet_by_index(0) 
        sheet1.cell_value(0, 0)
        
        notfileList.write(0,0,'S.No')
        notfileList.write(0,1,'Filename')
        notfileList.write(0,2,'Location')
        notfileList.set_column(1,1,30)
        notfileList.set_column(2,2,170)
        sno=0
        for i in range(1,count):
            diffMod=sheet.cell_value(i,2)
            diffCre=sheet.cell_value(i,3)
            extension=sheet1.cell_value(i,2)
            if(diffMod>=100 and (extension not in sample_extension)):
                sno+=1
                notfileList.write(sno,0,sno)
                filename=sheet1.cell_value(i,1)
                location=sheet1.cell_value(i,12)
                count_display.append(location[0])
                notfileList.write(sno,1,filename)
                notfileList.write(sno,2,location)
                
        notUsedfileList.close()
        file = r"NotUsedFile.xlsx"
        os.startfile(file)
        
        dic_display={}
        for i in range(0,len(count_display)):
            if count_display[i] not in dic_display:
                dic_display[count_display[i]]=count_display.count(count_display[i])

        x=[]
        count=[]
        for key,value in dic_display.items():
            x.append(key)
            count.append(value)

        plt.pie(count, labels=x,autopct='%1.1f%%',radius=5 ,shadow=True, startangle=140)
        plt.title("Percentage of Not used file in different drives")
        plt.axis('equal')
        plt.show()
        
                
            
        

class Scanning:
    def __init__(self):
        self.count=0
        fileList.write(0,0,'S.No')
        fileList.write(0,1,'File Name')
        fileList.write(0,2,'Extension')
        fileList.write(0,3,'Access time')
        fileList.write(0,4,'Modified time')
        fileList.write(0,5,'Created time')
        fileList.write(0,6,'Size(Bytes)')
        fileList.write(0,7,'Read Only')
        fileList.write(0,8,'Hidden')
        fileList.write(0,9,'System File')
        fileList.write(0,10,'Archive')
        fileList.write(0,11,'Compressed')
        fileList.write(0,12,'Location')
        fileList.set_column(1, 11, 30)
        fileList.set_column(12, 12, 170)
        
    def FilesInsert(self,filename,root):
        sample=os.path.join(root,filename)
        self.count=self.count+1
        fileList.write(self.count,0,self.count)
        fileList.write(self.count,1,filename)
        fileList.write(self.count,2,filename.split(".")[-1])
        fileList.write(self.count,3,time.ctime(os.path.getatime(sample)))
        fileList.write(self.count,4,time.ctime(os.path.getmtime(sample)))
        fileList.write(self.count,5,time.ctime(os.path.getctime(sample)))
        fileList.write(self.count,6,os.path.getsize(sample))
        
        file_flag = win32file.GetFileAttributesW(sample)
        is_readonly=file_flag & win32con.FILE_ATTRIBUTE_READONLY
        is_hidden = file_flag & win32con.FILE_ATTRIBUTE_HIDDEN
        is_system = file_flag & win32con.FILE_ATTRIBUTE_SYSTEM
        is_archive=file_flag & win32con.FILE_ATTRIBUTE_ARCHIVE
        is_comm=file_flag & win32con.FILE_ATTRIBUTE_COMPRESSED
        if(is_readonly==1):
                fileList.write(self.count,7,"Yes")
        else:
                fileList.write(self.count,7,"No")
        if(is_hidden==2):
                fileList.write(self.count,8,"Yes")
        else:
                fileList.write(self.count,8,"No")
        if(is_system==4):
                fileList.write(self.count,9,"Yes")
        else:
                fileList.write(self.count,9,"No")
        if(is_archive==32):
                fileList.write(self.count,10,"Yes")
        else:
                fileList.write(self.count,10,"No")
        if(is_comm==2048):
                fileList.write(self.count,11,"Yes")
        else:
                fileList.write(self.count,11,"No")
        fileList.write(self.count,12,root)
        return self.count
                
       
obj=Scanning()
print("Scanning process started")
available_drives = ["%s:" %d for d in string.ascii_uppercase if os.path.exists("%s:" %d)]
print("Available Drives:",available_drives)
i=0
start=time.time()
while(i<len(available_drives)):
    for root, dirs, files in os.walk(available_drives[i]):
        for filename in files:
            filescount=obj.FilesInsert(filename,root)
    i=i+1
end=time.time()
print("\nYour system has ",filescount," files\n")
print("Scanning process completed")
print("Scanning process time taken is",end-start)
objfileList.close()

obj1=Analysis()
obj1.DateDifference(filescount)
obj1.percentage(filescount)
year=obj1.year_percentage(filescount)
while(obj1.year_month_percentage(filescount,year)):
    dummy=0
    
    
obj1.duplicateFile(filescount)




print("Folderwise analysis started\n\n")
start=time.time()
folderfileList.write(0,0,'S.No')
folderfileList.write(0,1,'Location')
folderfileList.write(0,2,'Files count')
folderfileList.write(0,3,'Size(In Bytes)')
folderfileList.set_column(2, 3, 30)
folderfileList.set_column(1, 1, 115)
available_drives = ["%s:" %d for d in string.ascii_uppercase if os.path.exists("%s:" %d)]
j=0
dict_files={}
dict_size={}
folder_count=0
while(j<len(available_drives)):
    for root, dirs, files in os.walk(available_drives[j]):
        for i in range(0,len(dirs)):
            sample=os.path.join(root,dirs[i])
            size=os.path.getsize(root)
            for root1, dirs1, files1 in os.walk(sample):
                folder_count+=1
                folderfileList.write(folder_count,0,folder_count)
                folderfileList.write(folder_count,1,root1)
                folderfileList.write(folder_count,2,len(files1))
                dict_files.update({root1:len(files1)})
                size=0
                for file in files1:
                    sample1=os.path.join(root1,file)
                    size+=os.path.getsize(sample1)
                folderfileList.write(folder_count,3,size)
                dict_size.update({root1:size})
    j=j+1

sorted_files=sorted(dict_files.items(),key=operator.itemgetter(1))
sorted_size=sorted(dict_size.items(),key=operator.itemgetter(1))
fi=list(sorted_files[-1])
si=list(sorted_size[-1])
print("Folder which has more files is "+fi[0]+" with ",fi[1]," files")
print("Folder which consume more space is "+si[0]+" with ",si[1]," bytes\n\n")
end=time.time()
print("Folderwise process completed")
print("Folderwise process time taken is",end-start)
folderobjfileList.close()
anaobjfileList.close()


print("Not used file process started")
start=time.time()

obj1.notUsedFile(filescount)

end=time.time()
print("Not used file process completed")
print("Not used file process time taken is",end-start)



