import os
import sys
import json
import re
from datetime import datetime
from datetime import timedelta
import logging
 
logging.basicConfig(level=logging.WARNING, format=' %(levelname)s - %(lineno)d 行- %(message)s')
logger = logging.getLogger(__name__)

        
class GanZhi:
    def __init__(self):
        self.jqnames=["小寒", "大寒","立春", "雨水", "惊蛰", "春分", "清明", "谷雨", \
                        "立夏", "小满", "芒种", "夏至", "小暑", "大暑", \
                        "立秋", "处暑", "白露", "秋分", "寒露", "霜降", \
                        "立冬", "小雪", "大雪", "冬至"]
        self.tg=['甲','乙','丙','丁','戊','己','庚','辛','壬','癸']
        self.dz=['子','丑','寅','卯','辰','巳','午','未','申','酉','戌','亥']
        
        with open(os.path.join(os.path.dirname(__file__),'1800-2100jieqi.json'),'r',encoding='utf-8',errors='ignore') as f: 
            # for line in f.readlines():
            #     self.dic = json.loads(line)
            self.dic=json.load(f)

            # print(self.dic)

    
    def inputdate(self,y,m,d,h=-1,zishi=0):
        self.jqlist=self.dic[str(y)]
        
        ymd=str(y)+"-"+ str(m)+ '-' +str(d)
        self.jqdate=[]
        for i in self.jqnames:
            ptn=r'\d+年.+日'
            ptn_time=r'\d\d:\d\d:\d\d'
            ptn_lunar=r'农历.+'
            tmp_jq=re.search(ptn,self.jqlist[i])
            tmp_jq_time=re.search(ptn_time,self.jqlist[i])
            tmp_jq_lunar=re.search(ptn_lunar,self.jqlist[i])
            self.jqdate.append([i,tmp_jq[0],tmp_jq_time[0],tmp_jq_lunar[0]])
            
        logger.debug(self.jqdate) #当年的节气时间总表,数组形式
            
        for i in self.jqdate:
            i[1]=i[1].replace('年','-').replace('月','-')[0:-1]
            
        
        input_time=datetime.strptime(ymd, '%Y-%m-%d')
        self.jq_section=[]
        n=0

        for i in self.jqdate:
            list_time=datetime.strptime(i[1], '%Y-%m-%d')
            if list_time>input_time:
                self.jq_section.append(n-1)
                self.jq_section.append(self.jqdate[n-1][0])
                self.jq_section.append(self.jqdate[n][0])
                break            
            n+=1
        

        if self.jq_section:
            pass
        else:
            self.jq_section.append(23)
            self.jq_section.append('冬至')
            self.jq_section.append('小寒')            

            
        if (self.jq_section[0])%2==0:
            m_odr=self.jq_section[0]/2
        else:
            m_odr=(self.jq_section[0]-1)/2
            
        m_odr=m_odr if m_odr>0 else m_odr+12
            
            
        self.jq_section.append(m_odr)
        
        logger.debug(['输入的日期在今年的节气分段：',self.jq_section])  #[节气序数，节气，节气，月份的序号]
        
        return self.jq_section

    def cal_dateGZ(self,y,m,d,h=-1,zishi=0):
        
        
        
        # 因为23-0点涉及日期变动，先按输入参数校正
        # zishi==0, 23-0点按下一天的子时处理，默认
        # zishi==1, 23-0点按当天的子时处理
        _date=str(y)+'-'+str(m)+'-'+str(d)
        if h==23 and zishi==0:
            date_correct=datetime.strptime(_date,'%Y-%m-%d')+timedelta(days=1)     
            y=int(date_correct.strftime('%Y'))
            m=int(date_correct.strftime('%m'))
            d=int(date_correct.strftime('%d'))     
            
        dateGZ=self.inputdate(y,m,d)
        
        logger.info(['按子时校正后的日期：',y,m,d])
        
        #年干支
        odr_yg=int(str(y-3)[-1])-1
        odr_yz=int((y-3)%12)-1

        #月干支
#         odr_mg_0=self.month_gz_adjust(odr_yg)
      
        
        if dateGZ[3]==0:
            realMon=12
        else:
            realMon=dateGZ[3]
        
        
        odr_mg=int(self.gzodr((odr_yg+1)*2+realMon,'g'))%10-1
        odr_mz=int(self.gzodr(dateGZ[3]+1,'z'))  #1月是丑月，列表从子开始，故+1
        
        logger.debug(['月序数：',odr_mg,odr_mz,'未按节气校正后的年干支序数：', \
                      odr_yg,odr_yz,dateGZ[3],realMon,(odr_yg+1)*2+realMon])
        
        if m==1 and dateGZ[3]>10: #寅月前，算上一年
            odr_yg=int(str(y-1-3)[-1])-1

            odr_yz=odr_yz-1
            odr_mg=int(self.gzodr((odr_yg+1)*2+realMon,'g'))%10-1

            
        logger.debug(['按节气校正后的年干支序数：',odr_yg,odr_yz,odr_mg,dateGZ[3],realMon])
        
        #日干支
        #先计算这年的元旦干支,已知1900年元旦是 甲戌日
        this_yd=self.cal_yd(y)
        if this_yd[0]!=-1: #如果年份<1900，将返回-1
            days_diff=(datetime.strptime(str(y)+'-'+str(m)+'-'+str(d),'%Y-%m-%d')- \
                                            datetime.strptime(str(y)+'-1-1','%Y-%m-%d')).days
            odr_dg=(this_yd[0]+days_diff)%10
            odr_dz=(this_yd[1]+days_diff)%12
            
        else:
            pass       
        
        
        #时干支        
        if h!=-1:
            if h==23:
                h=0
                
            if h%2==0:  #按十二个时辰来计算
                H=h/2
            else:
                H=(h+1)/2    
                
            if odr_dg==0 or odr_dg==5:
                odr_hg_0=0
            elif odr_dg==1 or odr_dg==6:
                odr_hg_0=2
            elif odr_dg==2 or odr_dg==7:
                odr_hg_0=4
            elif odr_dg==3 or odr_dg==8:
                odr_hg_0=6
            elif odr_dg==4 or odr_dg==9:
                odr_hg_0=8

            odr_hg=int(odr_hg_0+H)
            odr_hz=int(H)
        
            res=[self.tg[self.gzodr(odr_yg,'g')], \
                    self.dz[self.gzodr(odr_yz,'z')], \
                    self.tg[self.gzodr(odr_mg,'g')], \
                    self.dz[self.gzodr(odr_mz,'z')], \
                    self.tg[self.gzodr(odr_dg,'g')], \
                    self.dz[self.gzodr(odr_dz,'z')], \
                    self.tg[self.gzodr(odr_hg,'g')], \
                    self.dz[self.gzodr(odr_hz,'z')] 
                ]
        else:
            res=[self.tg[self.gzodr(odr_yg,'g')], \
                    self.dz[self.gzodr(odr_yz,'z')], \
                    self.tg[self.gzodr(odr_mg,'g')], \
                    self.dz[self.gzodr(odr_mz,'z')], \
                    self.tg[self.gzodr(odr_dg,'g')], \
                    self.dz[self.gzodr(odr_dz,'z')], \
                    '', \
                    ''] 

        logger.debug(['输出结果:',res])
        logger.debug('----------------------------')
#         print(res)
        return res

    def gzodr(self,n,j): #校准超过10或12,或负数的天干地支序数
        if j=='g':
            if n>9:
                n=n-10
            if n<0:
                n=n+10
        elif j=='z':
            if n>11:
                n=n-12
            if n<0:
                n=n+12
        return n
    
    def cal_yd(self,n,o=1899): #先计算这年的元旦干支,已知1900年元旦是 甲戌日
        if n==1900:
            return [0,10]
        elif n>1900:            
            days=(datetime.strptime(str(n)+'-1-1','%Y-%m-%d')-datetime.strptime('1900-1-1','%Y-%m-%d')).days
            odr_tg=self.gzodr(days%10+0,'g')
            odr_dz=self.gzodr(days%12+10,'z')
            logger.debug(['距1900年元旦的天数：',days,'元旦的干支序数：',odr_tg,odr_dz])
            
            return [odr_tg,odr_dz]
        else:
            return [-1,-1]

    def dif_dates(self,n,o):
        d1=datetime.strptime(n,'%Y-%m-%d')
        d2=datetime.strptime(o,'%Y-%m-%d')
        return (d1-d2).days
          


if __name__=='__main__':
    mybz=GanZhi()
    
    k=mybz.cal_dateGZ(1975,2,14)
    
    print(k)