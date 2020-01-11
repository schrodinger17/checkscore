import requests
import re
import time
from bs4 import BeautifulSoup

payload = {'username': '**********', 'password': '**********'}
headers = {
    'User-Agent': 'Mozilla/5.0(Macintosh; Intel Mac OS X 10_11_4)\
        AppleWebKit/537.36(KHTML, like Gecko) Chrome/52 .0.2743. 116 Safari/537.36'
}

s = requests.session()
s.post('http://us.nwpu.edu.cn/eams/login.action', data=payload)
j = 0
for i in range(*******, ********):
    j += 1
    # print(j)

    if j == 30:
        s = requests.session()
        s.post('http://us.nwpu.edu.cn/eams/login.action', data=payload)
        j = 0
    response = s.get('http://us.nwpu.edu.cn/eams/teach/grade/course/person!info.action?courseGrade.id=' + str(i),
                     headers=headers)
    # print(response.text)
    # print("\n\n\n\n")
    content = response.text
    # result=re.match(r'^<td class="title">得分</td>(.*?)(\d)</span></td>$',content,re.I|re.M)
    soup = BeautifulSoup(content, 'lxml')
    print("正在查询" + " " + str(i))
    find1 = False
    find2 = False
    find3 = False
    find4 = False
    find5 = False
    for table in soup.find_all(name='table'):
        for td in table.find_all(name='td'):
            # print(td.attrs)
            td.attrs.setdefault('class', '')
            # print(td.string=='***')
            if not find1 and td.attrs['class'] == ['title']:
                if td.string == 'Name':
                    find1 = True
                    continue
            if not find2 and find1 and td.attrs['class'] == ['content']:
                if td.string == '***':
                    # stu=td.string
                    find2 = True
                    # print('find')
                    continue
            if not find3 and find2 and td.attrs['class'] == ['title']:
                if td.string == 'Course Name':
                    find5 = True
                    continue
            if find5 and td.attrs['class'] == ['content']:
                coursename = td.string
                find5 = False
                continue
            if not find3 and find2 and td.attrs['class'] == ['title']:
                if td.string == '得分':
                    find4 = True
                    continue
            if find4 and td.attrs['class'] == ['content']:
                f = open("test.txt", "a+")
                print(str(i) + ' ' + str(coursename) + ':' + str(td.string) + '\n', file=f)
                f.close()
                # print(str(stu)+'   '+str(coursename)+':'+str(td.string))
                break
    # print(result)
    # s.get('http://us.nwpu.edu.cn/eams/home.action')
    time.sleep(0.1824465)

