import re
import xlsxwriter
import sys
import os

#GLOBAL VARIABLES:
filepath=""
filename=""
try:
    filepath, filename = os.path.split(sys.argv[1])
except:
    print "You have to drop an HTML file of your timetable on top of this script in order for it to work"
    raw_input()
    sys.exit()
DEBUGMODE=False
conflicts={};
coursename=""
courseweeks=""
WEEKDAYS_COUNT=5
IGNORE_FIELDS_OFFSET=7
WEEKDAY_NAMES={"MONDAY":0,"TUESDAY":1,"WEDNESDAY":2,"THURSDAY":3,"FRIDAY":4,"SATURDAY":5,"SUNDAY":6};
MONTH_LIST=["SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER","JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE"]
MONTH_LENGTHS={MONTH_LIST[0]:30,MONTH_LIST[1]:31,MONTH_LIST[2]:30,MONTH_LIST[3]:31,MONTH_LIST[4]:31,
               MONTH_LIST[5]:28,MONTH_LIST[6]:31,MONTH_LIST[7]:30,MONTH_LIST[8]:31,MONTH_LIST[9]:30};
WEEKDAY_COUNT=7
#EXCEL STYLES:
EXCEL_MONTH_MARGINTOP=4
EXCEL_MONTH_MARGINBOTTOM=2
EXCEL_HOUR_MARGINLEFT=0
EXCEL_HOUR_START=8
EXCEL_HOUR_END=21
EXCEL_HOUR_INTERVAL=0.5
EXCEL_1STCOL_WIDTH=20
EXCEL_DAY_ROWHEIGHT=60
EXCEL_DAY_COLWIDTH=8
#CLASSES VARIABLES:
CLASSES_STARTWEEK_OFFSET=8
CLASSES_MODULE_TABLE_ELEM_COUNT=7
CLASSES_MODULE_MAX_COUNT=10

def list_indexof(lst,str):
    for i in range(len(lst)):
        if lst[i]==str:
            return i
    return -1

def dict_sumvals(dict,length):
    s=0
    for i in range(length):
        s+=dict[MONTH_LIST[i]]+EXCEL_MONTH_MARGINBOTTOM+1
    return s

def drange(start,stop,step):
    r=start
    while r<=stop:
        yield r
        r+=step

def dict_getkey_byval(dict,value):
    for key,val in dict.iteritems():
        if val==value:
            return key
                
def create_timetable_structure(filepath,filename):
    global coursename,courseweeks
    timetable_struct=[]
    timetable_html_file=open(filepath+"\\"+filename,"r")
    print filepath+"\\"+filename
    timetable_html_text=timetable_html_file.read()
    #FETCH ALL WEEK DAYS
    match_weekdays=re.findall(r'<p>(?:.|\n)+?labelone(?:.|\n)+?table>', timetable_html_text, re.S|re.M)
    if match_weekdays:
        for week_ctr in range(WEEKDAYS_COUNT):
            table_match=re.findall(r'(?:<td>(.+?)<\/td>)',match_weekdays[week_ctr],0)
            table_match=table_match[IGNORE_FIELDS_OFFSET:]
            metadata=re.findall(r'Course:.+?\">(.+?)<.+?Weeks:.+?">(.+?)<.+?\(.+?">(.+?)<',timetable_html_text,re.S)
            coursename=metadata[0][0]
            courseweeks="Weeks: "+metadata[0][1]+" ("+metadata[0][2]+")"
            
            #FIX WEEKS IN THE MATCH:
            weeks_splitted=[]
            weeks_index=[]
            for k in range(len(table_match)):
                if table_match[k-1].find(":")>0 and table_match[k-2].find(":")>0:
                    weeks_splitted.append(table_match[k].split(","))
                    weeks_index.append(k)
                    
            for k in range(len(weeks_splitted)):
                for j in range(len(weeks_splitted[k])):
                    if(weeks_splitted[k][j].find("-")>0):
                        weeks_splitted[k][j]=weeks_splitted[k][j].split("-")
                        
            for k in range(len(weeks_index)):
                table_match[weeks_index[k]]=weeks_splitted[k]
            table_match=[dict_getkey_byval(WEEKDAY_NAMES,week_ctr)]+table_match
            #WEEKS FIXED: DONE
            timetable_struct.append(table_match)
    else:
        print "Weekdays couldn't be matched"
    timetable_html_file.close()
    return timetable_struct

def get_days_byweeks(weekday,days_interval):
    day_timetable_start=CLASSES_STARTWEEK_OFFSET*WEEKDAY_COUNT
    week_start=0
    using_lists=False
    if(isinstance(days_interval,list)):
        week_start=int(days_interval[0])
        using_lists=True
    else:  
        week_start=int(days_interval)
    
    for month in range(len(MONTH_LIST)):
        for days in range(MONTH_LENGTHS[MONTH_LIST[month]]):
            day_timetable_start+=1
            if day_timetable_start==week_start*7+2:
                daylist=[]
                if(using_lists):
                    for i in range(int(days_interval[1])-week_start):
                        if (days+(i*7)+WEEKDAY_NAMES[weekday])>MONTH_LENGTHS[MONTH_LIST[month]]:
                            days-=MONTH_LENGTHS[MONTH_LIST[month]]
                            month+=1
                        daylist.append([MONTH_LIST[month],(days+(i*7)+WEEKDAY_NAMES[weekday])])
                else:
                    daylist.append([MONTH_LIST[month],(days+WEEKDAY_NAMES[weekday])])
                return daylist  



def cap(s, l):
    return s if len(s)<=l else s[0:l-3]+'...'

def create_excelfile(timetable,filepath):
    global coursename,courseweeks
    workbk=xlsxwriter.Workbook(filepath+"\\course_"+coursename.lower().replace(" ","_").replace("/","_")+"_timetable.xlsx")
    worksht=workbk.add_worksheet(cap(coursename,21)+" timetable")
    #EXCEL STYLES:
    EXCEL_STYLE1=workbk.add_format({'bold':True,'bg_color':"#1C6D73",'font_size':20,'font_color':"#FFFFFF",'align':'center'}) # MONTHS
    EXCEL_STYLE2=workbk.add_format({'bold':True,'bg_color':"#68B7BD",'align':'center','valign':'vcenter'}) # HOURS
    EXCEL_STYLE3=workbk.add_format({'bold':True,'bg_color':"#68B7BD",'align':'center','valign':'vcenter','text_wrap':True}) # DAYS
    EXCEL_STYLE4=workbk.add_format({'bg_color':'#CADDDE','valign':'top','align':'left','text_wrap':True,'border':1,'font_size':9}) # CLASSES
    EXCEL_STYLE5=workbk.add_format({'bg_color':'#133E69','bold':True,'font_size':18,'border':1,'valign':'vcenter','text_wrap':True,'align':'center','font_color':'#FFFFFF'}); # TITLE
    EXCEL_STYLE6=workbk.add_format({'bg_color':'#56799C','border':1,'valign':'vcenter','text_wrap':True,'align':'center','font_color':'#FFFFFF'}); #WEEKS
    # CONSTRUCT EXCEL FILE - BEGIN 
    
    #INSERT STRUCTURE OF MONTH, DAYS AND HOURS:
    worksht.set_column("A:A",EXCEL_1STCOL_WIDTH)
    #TITLE, YEAR AND WEEKS:
    worksht.set_column("B:AB",EXCEL_DAY_COLWIDTH)
    worksht.merge_range('A1:H3',"Course: "+coursename,EXCEL_STYLE5)
    worksht.merge_range('I1:K3',courseweeks,EXCEL_STYLE6)
    #REST OF THE STRUCTURE:
    day_offset=EXCEL_MONTH_MARGINTOP
    day_permctr=0
    weekday_ctr=-1
    for month in range(len(MONTH_LIST)):
        hour_ctr=EXCEL_HOUR_MARGINLEFT
        #MONTHS:
        worksht.write(day_offset,0,MONTH_LIST[month],EXCEL_STYLE1)
        #HOURS:
        for hour in drange(EXCEL_HOUR_START,EXCEL_HOUR_END,EXCEL_HOUR_INTERVAL):
            worksht.write(day_offset,hour_ctr+1,str(int(hour))+":%02dh"%int((hour%1)*60),EXCEL_STYLE2)
            hour_ctr+=1
        #DAYS AND WEEK COUNTER:
        for day in range(MONTH_LENGTHS[MONTH_LIST[month]]):
            worksht.set_row(day+day_offset+1,EXCEL_DAY_ROWHEIGHT)
            week_msg=""
            if((day+day_permctr)%WEEKDAY_COUNT==0):
                week_msg="(WEEK "+str((day+day_permctr)/WEEKDAY_COUNT+CLASSES_STARTWEEK_OFFSET)+")\n"
            if(weekday_ctr>5):
                weekday_ctr=0
            else:
                weekday_ctr+=1
            worksht.write(day+day_offset+1,0,week_msg+dict_getkey_byval(WEEKDAY_NAMES, weekday_ctr)+"\nDay "+str(day+1)+":",EXCEL_STYLE3)
            
        day_permctr+=day+1
        day_offset=day_offset+day+EXCEL_MONTH_MARGINBOTTOM+2
 
    #INSERT DATA INTO THE EXCEL FILE
    timetable_week_index=4
    timetable_weeksfixed=[]
    for weekday in range(len(timetable)):
        timetable_noweekday=timetable[weekday][1:]
        
        if DEBUGMODE:
            print dict_getkey_byval(WEEKDAY_NAMES, weekday),":"
        for module_ctr in range(CLASSES_MODULE_MAX_COUNT):
            module_start=module_ctr*CLASSES_MODULE_TABLE_ELEM_COUNT
            module_end=module_ctr*CLASSES_MODULE_TABLE_ELEM_COUNT+CLASSES_MODULE_TABLE_ELEM_COUNT
            
            if(timetable_noweekday[module_start:module_end]):
                weekday_list=timetable_noweekday[module_start:module_end][timetable_week_index]
                daylist=[]
                for k in range(len(weekday_list)):
                    daylist.append(get_days_byweeks(dict_getkey_byval(WEEKDAY_NAMES, weekday),weekday_list[k]))
                module=timetable_noweekday[module_start:module_end]
                module[timetable_week_index]=daylist
                timetable_weeksfixed.append(module)
                
                if DEBUGMODE:
                    print "Module ",module_ctr,module
                hour_begin=re.findall(r"([0-9].+?)",module[2])
                hour_end=re.findall(r"([0-9].+?)",module[3])
                class_col_begin=(int(hour_begin[0])-8)*2+1
                class_col_end=(int(hour_end[0])-8)*2+1
                if(hour_begin[1].find("30")>-1):
                    class_col_begin+=1
                if(hour_end[1].find("30")>-1):
                    class_col_end+=1
                class_rows=[]
                 
                    
                for i in range(len(module[timetable_week_index])):
                    for j in range(len(module[timetable_week_index][i])):
                        monthname=module[timetable_week_index][i][j][0]
                        monthindex=list_indexof(MONTH_LIST,monthname)
                        day_monthoffset=module[timetable_week_index][i][j][1]
                        class_rows.append(dict_sumvals(MONTH_LENGTHS, monthindex)+EXCEL_MONTH_MARGINTOP+day_monthoffset)
                
                #INSERT CLASSES:
                for i in range(len(class_rows)):
                    if(class_rows[i] in conflicts):
                        if class_col_begin==conflicts[class_rows[i]][1]:
                            class_col_begin+=1
                            class_col_end+=1
                        else:
                            if(class_col_begin>conflicts[class_rows[i]][0] and class_col_end<=conflicts[class_rows[i]][1]):
                                break
                        if DEBUGMODE:
                            print "Conflict: Row: ",class_rows[i]," Col begin: ",class_col_begin,",",conflicts[class_rows[i]][0]," Col end: ",class_col_end,",",conflicts[class_rows[i]][1]
                    
                    #CLASS WAS ADDED TO A DICTIONARY TO PREVENT FUTURE CONFLICTS
                    conflicts[class_rows[i]]=[class_col_begin,class_col_end]
                    #INSERT CLASS:
                    worksht.merge_range(class_rows[i],class_col_begin,class_rows[i],class_col_end,
                                        module[0]+"\n"+module[1]+"\n"+module[2]+"h-"+module[3]+"h\n"+
                                        module[5]+" "+module[6],EXCEL_STYLE4)
        print ""
    # CONSTRUCT EXCEL FILE - END
    workbk.close()
    
def main(filepath,filename):
    create_excelfile(create_timetable_structure(filepath, filename),filepath)  
main(filepath,filename)