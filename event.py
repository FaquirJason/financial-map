import xlrd
import re 
import csv

class Event(object):
   def __init__(self, uid, time, result_type, result_country, content, emotion,title,source,sentence,trigger):
       self.uid = uid  
       self.time = time
       self.result_type = result_type
       self.result_country = result_country
       self.content = content
       self.emotion = emotion
       self.title = title
       self.source = source
       self.sentence = sentence
       self.trigger = trigger

   

def read_data(pas,file):
    data = xlrd.open_workbook(pas+"/"+file)
    table = data.sheets()[0]
    nrows = table.nrows 

    events = []

    for i in range(1,nrows):
        row_detail = table.row_values(i)
        row_detail[5] = row_detail[5][1:-1]
        row_detail[5] = re.sub(r'\s+', "", row_detail[5])
        row_detail[5] = re.sub(r'\[', ' ', str(row_detail[5]))
        row_detail[5] = re.sub(r'\]', '+', str(row_detail[5]))
        row_detail[5] = re.sub(r'\,', ' ', str(row_detail[5]))
        row_detail[5] = re.sub(r'\'', ' ', str(row_detail[5]))
        row_detail[4] = re.sub(r'\#+', '。', str(row_detail[4]))
        row_detail[4] = re.sub(r'\：', '。', str(row_detail[4]))
        row_detail[5] = row_detail[5].split()
        j = 0
        # event = Event(row_detail[0],row_detail[1],row_detail[5][j+1],row_detail[5][j],row_detail[4],row_detail[5][j+2],row_detail[3],row_detail[2])
        # print(event.result_type)
        while j < len(row_detail[5]):

            event = Event(row_detail[0],row_detail[1],row_detail[5][j+1],row_detail[5][j],row_detail[4],row_detail[5][j+2],row_detail[3],row_detail[2],None,None)
            j += 4
            events.append(event)

    return events

def read_dictionary(pas,file):
    data = xlrd.open_workbook(pas+"/"+file)
    table = data.sheets()[0]
    nrows = table.nrows 

    trigger = []

    for i in range(1,nrows):
        row_detail = table.row_values(i)
        trigger.append(row_detail[0])
        trigger.append(row_detail[1])
        trigger.append(row_detail[2])
    
    while '' in trigger:
        trigger.remove("")

    return trigger


if __name__ == '__main__':
    events = read_data(".","舆情事件及情感识别.xlsx")
    # for event in events:
    #     print(event.uid)

    trigger = read_dictionary(".","金融术语-经济指标.xlsx")
    # trigger = trigger[0:5]
    read_trigger = "|".join(trigger)

    for i in range(0,len(events)):
        sentences = re.findall('([^。]*('+read_trigger+')[^。]*)',events[i].content)
        for item in sentences:
            events[i].sentence = item[0]
            events[i].trigger = item[1]


    with open("Triples_Event.csv",'w') as f:
        csv_write = csv.writer(f)
        for event in events:
            if event.trigger != None:
                csv_write.writerow(["V","trigger",event.trigger,event.trigger])
                # print(event.trigger)
                csv_write.writerow(["V","event",event.uid,event.sentence])
                # print(event.sentence)
                csv_write.writerow(["V","event_type",event.result_type,event.result_type])
                csv_write.writerow(["E","包含",event.uid,event.trigger])
                csv_write.writerow(["E","属于",event.uid,event.result_type])

    


