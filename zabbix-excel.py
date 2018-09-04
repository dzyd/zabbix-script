#coding=utf8
import time
import datetime
from urllib import request,parse
import json
import xlrd
import xlwt

#a week 604800s
apiurl="http://your-zabbix-address/api_jsonrpc.php"
username="xxx"
password="xxx"
hostids=[]
hostnm=[]
#指定时间的代码
'''
htime = "2018-08-27 00:00:00"
xtime = "2018-09-03 00:00:00"
ss=time.strptime(xtime,"%Y-%m-%d %H:%M:%S")
print(ss)
ans_time=time.mktime(ss)
print(ans_time)
an_time=ans_time-604800
an_time=ans_time-1800
show_time=htime + '-' + xtime
#自动获取时间的代码
'''
nowTime=datetime.datetime.now()
nowime=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
htime=(nowTime - datetime.timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S')
show_time=htime + '-' + nowime

ans_time = time.mktime(nowTime.timetuple())
an_time=ans_time-604800


def http_post(values):
    #print (values)
    HEADERS = {"Content-Type": "application/json"}
    jdata = json.dumps(values).encode('utf-8')             # 对数据进行JSON格式化编码
    print (jdata)
    req = request.Request(apiurl, data=jdata,headers=HEADERS)
    re = request.urlopen(req)
    html = re.read().decode('utf-8')
    return html

def gettoken():
    usr_value ={"jsonrpc": "2.0","method": "user.login","params": {"user": username,"password": password},"id": 1}
    request = http_post(usr_value)
    page = format(json.loads(request).get('result'))
    token = page
    print("token is:",token)
    return token

def gethostid(token):
    global hostids
    global hostnm
    host_value={
"jsonrpc": "2.0",
"method": "hostgroup.get",
"params": {
    "output": "extend",
    "selectHosts":"extend",
    "filter":{"name":"online"}
},
"auth": token,
"id": 1
}

    request = http_post(host_value)
    page = json.loads(request).get('result')
    #print('sss',hostids)

    for j in page[0]["hosts"]:
        #print(j["host"],j["hostid"])
        hostids.append(j['hostid'])  #get IDs for host
        hostnm.append(j['host'])  # get names for host
    print("online group member",hostnm)
    #print(hostids)
    #print(hostnm)'
#在zabbix建一个week应用集，将需要监控的项目丢进去
def getitems(hid):
    itemnms=[]
    itemvalues=[]
    item_value={
    "jsonrpc": "2.0",
    "method": "application.get",
    "params": {
        "output": ["name"],
        "hostids": hid,
        "selectItems":"extend",
        "filter":{"name":"week"}
        #"sortfield": "name"
    },
    "auth": token,
    "id": 1
}

    request = http_post(item_value)
    page = json.loads(request).get('result')
    #print ("all item data",page)
    dd={}
    for i in page[0]['items']:
        print ("the currect item is:",i["key_"])
        itemnms.append(i['key_'])#item的名字
        itemid=getitemid(i['key_'],hid)
        if  i["key_"].find('cpu')>0:
            historyid=0
        #if  i["key_"].find('cpu')>0:
        #    historyid=0
        else:
            historyid=3
        itemvalue=getvalues(itemid,historyid)#item的值
        itemvalues.append(itemvalue)

    #print("itemnms %s  itemvalues %s"%(itemnms,itemvalues))
    dd=dict(zip(itemnms,itemvalues))

    print("all the data is",dd)
    #print("dd[0]",dd['system.cpu.load[all,avg5]'])
    #print(page[0]['items'][1])
    #print(page[0]['items'][0]['key_'])
    return dd
def getitemid(itemnm,hostid):
    itemid_value={"jsonrpc":"2.0","method":"item.get","params":{"output":["itemid"],"hostids":hostid,"search":{"key_":itemnm}},"auth":token,"id":1}
    request = http_post(itemid_value)
    page = json.loads(request).get('result')
    print ("this is %s itemid %s"%(hostid,page[0]['itemid']))
    return page[0]['itemid'] #返回item的id

def getvalues(id,historyid):
    num=0
    value_value={"jsonrpc": "2.0","method": "history.get","params": {"output": "extend","history": historyid,"itemids": id,"time_from": an_time,"time_till": ans_time},"id": 1,"auth":token}
    request = http_post(value_value)
    page = json.loads(request).get('result')
    #print("all values return",page)
    for valuen in page:
        #print(valuen)
        num=float(valuen['value'])+float(num)
        #print(valuen['value'],num)
    if historyid==0:
        #print("all time total",num)
        avg=num/len(page)
        #print("the avg value is:",avg)
    else:
        #print("all time total",num)
        avg=num/len(page)/1024/1024/1024
        #print("the avg value is:",avg)
    return avg

def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    patterni= xlwt.Pattern()
    patterni.pattern=1
    patterni.pattern_fore_colour=3 #背景色3为绿色
    font = xlwt.Font()  #创建一个文本格式，包括字体、字号和颜色样式特性
    font.name = name
    font.bold = bold
    font.color_index = 4 #设置颜色为4
    font.height = height
    style.font = font
    style.pattern=patterni
    return style

def data_style(color):
    style = xlwt.XFStyle()
    patterni= xlwt.Pattern()
    patterni.pattern=1
    patterni.pattern_fore_colour=color
    style.pattern=patterni
    return style

def write_excel(excel_data,hostidd,a):
    odata = xlrd.open_workbook('demo1.xls');
    otable = odata.sheets()[0]
    onrows = otable.nrows
    ocolumn = otable.ncols
    print(ocolumn)
    print (onrows)
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1')
    row = [u'时间',u'服务器',u'当前负载',u'硬盘状态',u'OPT分区',u'内存']
    row0 = [u"ww"]
    row1 = [u'load1',u'load5',u'load15',u'总大小',u'已用',u'总大小',u'已用',u'总大小',u'已用']
    u=0
    for u in range(0,len(row)):
        print(row[u],u)
        if u==0:
            sheet1.write_merge(0,1,0,0,row[u],set_style('Arial',220,True))
        if u == 1:
            sheet1.write_merge(0,1,1,1,row[u],set_style('Arial',220,True))
        if u==2:
            sheet1.write_merge(0,0,2,4,row[u],set_style('Arial',220,True))
        if u==3:
            sheet1.write_merge(0,0,5,6,row[u],set_style('Arial',220,True))
        if u==4:
            sheet1.write_merge(0,0,7,8,row[u],set_style('Arial',220,True))
        if u==5:
        #else:
            sheet1.write_merge(0,0,9,10,row[u],set_style('Arial',220,True))
    for j in range(0,len(row1)):
        sheet1.write(1,j+2,row1[j],set_style('Arial',220,True))
    color =44
    for d in range(2, onrows):
        if d%10>=2:
            if (d//10)%2==0:
                color=44
            else:
                color=22
        if d%10<2:
            if (d//10)%2==0:
                color=22
            else:
                color=44

        if d%10==2:
            value = otable.cell(d,0).value
            sheet1.write_merge(d,d+9,0,0,value,data_style(color))
            for e in range(1,ocolumn):
                value = otable.cell(d,e).value
                sheet1.write(d,e,value,data_style(color))
        else:
            for e in range(1,ocolumn):
                value = otable.cell(d,e).value
                sheet1.write(d,e,value,data_style(color))
    sheet1.write(onrows,1,hostidd) #插入主机
    if a==0:
        sheet1.write_merge(onrows,onrows+9,0,0,show_time)  #插入时间
    for i in excel_data:
        if i.find('load')>0:
            if i.find('avg5')>=0:
                print('load54',i)
                sheet1.write(onrows,3,excel_data[i],data_style(color))
                continue
                print('success')
            if i.find('avg15')>=0:
                sheet1.write(onrows,4,excel_data[i],data_style(color))
                continue
                #print('success')
            if i.find('avg1')>=0:
                print("xxxxxxxxx",excel_data[i],data_style(color))

                #print("xxxxxxxxx",i)
                sheet1.write(onrows,2,excel_data[i],data_style(color))
                continue
                #print('success')
        if i.find('memory')>0:
            if i.find("total")>0:
                sheet1.write(onrows,9,excel_data[i],data_style(color))
                continue
                print('memory',i)
            if i.find("available")>0:
                sheet1.write(onrows,10,excel_data[i],data_style(color))
                continue
        if i.find('vfs')>=0:
            if i.find('project')>0 or i.find('data')>0:
                if i.find('total')>0:
                    sheet1.write(onrows,11,excel_data[i],data_style(color))
                    continue
                    #print("vfs",i)
                if i.find('used')>0:
                    sheet1.write(onrows,12,excel_data[i],data_style(color))
                    continue
            if i.find("opt")>0:
                if i.find("total")>0:
                    sheet1.write(onrows,7,excel_data[i],data_style(color))
                    continue
                if i.find('used')>0:
                    sheet1.write(onrows,8,excel_data[i],data_style(color))
                    continue
            else:
                if i.find("total")>0:
                    sheet1.write(onrows,5,excel_data[i],data_style(color))
                    continue
                if i.find('used')>0:
                    sheet1.write(onrows,6,excel_data[i],data_style(color))
                    continue
    f.save('demo1.xls')




token=gettoken()
gethostid(token)
#for hid in hostids:
for hid in range(0,len(hostids)):
    print("the currect host is",hostnm[hid])
    allitems=getitems(hostids[hid]) #获得一个主机的全部数据
    write_excel(allitems,hostnm[hid],hid)
