import xlwings as xw
import json
from src.lib import message, storage
import pandas
from env import SEND


def search_phone_num_student(name):
        student_num_list = [] 
        with open('sdb_student_db.csv',encoding='utf-8') as f:
            sdb_db = f.read()
            
        data_list = [line.split(',') for line in sdb_db.split('\n')]
        data_list[0][0] = data_list[0][0].lstrip('\ufeff')
        
        for p in range(len(data_list)):
            if data_list[p][0] == name:
                #학생이름을 통해 학생 전화번호 찾기
                #print(data_list[p][3])
                student_num_list.append(data_list[p][3])
            
        print("TEST", student_num_list)
        
        return student_num_list

    
def search_phone_num_parents(name):
        parents_num_list = [] 
        with open('sdb_student_db.csv',encoding='utf-8') as f:
            sdb_db_parents = f.read()
            
        parents_data_list = [line.split(',') for line in sdb_db_parents.split('\n')]
        
        for i in range(len(parents_data_list)):
            if parents_data_list[i][0] == name:
                #학생이름을 통해 부모님 전화번호 찾기
                parents_num_list.append(parents_data_list[i][4])
            
        #print(self.send_num_parents)
        
        return parents_num_list
    

def main():
    # 엑셀 파일 열기
    wb = xw.Book('main.xlsm')

    # 세 번째 시트 선택 (인덱스는 0부터 시작하므로 2번째 시트는 인덱스 1)
    sheet = wb.sheets[0]
    sheet_final = wb.sheets[1]
    i=1 
    #CDEFGHIJKL
    used_range = sheet.used_range
    num = sheet.range('C18').value
    
    for p in range(1,int(num)+1):
        row_values = used_range.rows[p].value
        written_message = f'이름 : {row_values[2]} \n출결 : {row_values[3]} \n진도 : {row_values[8]} \n과제수행도 : {row_values[4]} \n플래너수행도 : {row_values[5]} \n벌점 : {row_values[6]} \n테스트 결과 : {row_values[7]}'
        
        written_message_name = row_values[2]

        sheet_final[f'C{p+1}'].value = written_message 
        sheet_final[f'A{p+1}'].value = written_message_name
        
        #전송 전화번호 호출
        student_send_number = search_phone_num_student(written_message_name)
        parents_send_number = search_phone_num_parents(written_message_name)

        print("전송할 번호", student_send_number)

        data = {
                'messages': [
                    {
                        'to': student_send_number,
                        'from': SEND.SENDNUMBER,
                        'subject': '수다방학원',
                        'text': written_message
                    }
                ]
            }
            
        res = message.send_many(data)
        print(f"{student_send_number}에게 성공적으로 전송했습니다")
        print(json.dumps(json.loads(res.text), indent=2, ensure_ascii=False))
        
        
        data2 = {
                'messages': [
                    {
                        'to': parents_send_number,
                        'from': SEND.SENDNUMBER,
                        'subject': '수다방학원',
                        'text': written_message
                    }
                ]
            }
            
        res = message.send_many(data2)
        print(f"{student_send_number}에게 성공적으로 전송했습니다")
        print(json.dumps(json.loads(res.text), indent=2, ensure_ascii=False))
        
        
        
        
        
    


    


if  __name__== "__main__":
    main()

