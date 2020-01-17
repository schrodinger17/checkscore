# -*- coding:gbk -*-

import requests
from bs4 import BeautifulSoup
import xlwt as wt


class Score(object):
    def __init__(self, regular, experiment, final, total):
        self.regular = regular
        self.experiment = experiment
        self.final = final
        self.total = total


class Course(object):
    def __init__(self, name, semester, credit, score):
        self.name = name
        self.semester = semester
        self.credit = credit
        self.score = score


payload = {'username': '******', 'password': '******'}
headers = {
    'User-Agent': 'Mozilla/5.0(Macintosh; Intel Mac OS X 10_11_4)\
        AppleWebKit/537.36(KHTML, like Gecko) Chrome/52 .0.2743. 116 Safari/537.36'
}

s = requests.session()
s.post('http://us.nwpu.edu.cn/eams/login.action', data=payload, headers=headers)

response = s.get('http://us.nwpu.edu.cn/eams/teach/grade/course/person!historyCourseGrade.action?projectType=MAJOR',
                 headers=headers)
response.encoding = 'utf-8'
with open('test.html', 'w+') as f:
    # print(response.text)
    f.write(response.text)
content = response.text
soup = BeautifulSoup(content, 'lxml')
li = []
courses = []
for tbody in soup.find_all(name='tbody'):
    # print(tbody.attrs)
    tbody.attrs.setdefault('id', '')
    if tbody.attrs['id'] != '':
        for td in tbody.find_all(name='td'):
            # print(td)
            li.append(td)
            # print(len(li))
            if len(li) == 13:
                # print(li)
                # print(li[3].a.string)
                course = Course(li[3].a.string, li[0].string, li[5].string,
                                Score(li[6].string, li[7].string, li[8].string, li[10].string))
                courses.append(course)
                li.clear()
    else:
        continue
# print(courses)
workbook = wt.Workbook(encoding='utf-8')
sheet_names = []
sheets = []
lines = []
max_widths = []
sums = []
credit_sums = []
style = wt.XFStyle()
font = wt.Font()  # 为样式创建字体
font.name = '微软雅黑'
font.height = 0x00C8  # 默认高度
font.bold = False  # 粗体
font.underline = False  # 下划线
font.italic = False  # 斜体
style.font = font  # 设定样式
alignment = wt.Alignment()
alignment.vert = alignment.VERT_CENTER  # 垂直居中
alignment.horz = alignment.HORZ_CENTER  # 水平居中
# alignment.wrap = alignment.WRAP_AT_RIGHT  # 自动换行
style.alignment = alignment
for i in range(len(courses)):
    if courses[i].semester not in sheet_names:
        sheet = workbook.add_sheet(courses[i].semester)
        sheet.col(0).width = 256 * 2 * 5
        font.bold = True
        style.font = font
        sheet.write(0, 0, '科目', style)
        sheet.write(0, 1, '学分', style)
        sheet.write(0, 2, '平时成绩', style)
        sheet.write(0, 3, '实验成绩', style)
        sheet.write(0, 4, '期末成绩', style)
        sheet.write(0, 5, '总评成绩', style)
        sheet_names.append(courses[i].semester)
        sheets.append(sheet)
        lines.append(1)
        max_widths.append(5)
        sums.append(0)
        credit_sums.append(0)
    font.bold = False
    style.font = font
    sheet = sheets[sheet_names.index(courses[i].semester)]
    line = lines[sheet_names.index(courses[i].semester)]
    max_width = max_widths[sheet_names.index(courses[i].semester)]
    score_sum = sums[sheet_names.index(courses[i].semester)]
    credit_sum = credit_sums[sheet_names.index(courses[i].semester)]
    width = len(courses[i].name)
    if width > max_width:
        max_width = width
        max_widths[sheet_names.index(courses[i].semester)] = max_width
        sheet.col(0).width = 256 * 2 * max_width
    sheet.write(line, 0, courses[i].name, style)
    sheet.write(line, 1, courses[i].credit, style)
    sheet.write(line, 2, courses[i].score.regular, style)
    sheet.write(line, 3, courses[i].score.experiment, style)
    sheet.write(line, 4, courses[i].score.final, style)
    sheet.write(line, 5, courses[i].score.total, style)
    if 'P' not in courses[i].score.total:
        score_sum += float(courses[i].score.total) * float(courses[i].credit)
        credit_sum += float(courses[i].credit)
    line += 1
    lines[sheet_names.index(courses[i].semester)] = line
    sums[sheet_names.index(courses[i].semester)] = score_sum
    credit_sums[sheet_names.index(courses[i].semester)] = credit_sum
sheet0 = workbook.add_sheet('总成绩')
max_width = 0
for i in sheet_names:
    width = len(i)
    if width > max_width:
        max_width = width

for i in range(len(sheet_names)):
    sheet = sheets[i]
    line = lines[i]
    score_sum = sums[i]
    credit_sum = credit_sums[i]
    average_score = round(score_sum / credit_sum, 2)
    font.bold = True
    style.font = font
    sheet.write(line + 1, 0, '学分绩', style)
    sheet0.write(i, 0, sheet_names[i], style)
    sheet0.col(0).width = 256 * 2 * max_width
    font.bold = False
    style.font = font
    sheet.write(line + 1, 1, average_score, style)
    sheet0.write(i, 1, average_score, style)

score_sum = 0
credit_sum = 0
for i in sums:
    score_sum += i
for i in credit_sums:
    credit_sum += i
average_score = round(score_sum / credit_sum, 2)
font.bold = True
style.font = font
sheet0.write(len(sheet_names), 0, '总学分绩', style)
font.bold = False
style.font = font
sheet0.write(len(sheet_names), 1, average_score, style)
workbook.save("成绩.xlsx")
