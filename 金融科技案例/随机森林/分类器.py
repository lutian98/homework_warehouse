import xlwt
import xlrd

path = '作业：信用卡风险识别数据.xlsx'

data = xlrd.open_workbook(path)
table = data.sheet_by_index(0)  # sheet1
rowsnum = table.nrows  # 行数
colsnum = table.ncols  # 列数
rowsname = table.row_values(0)  # 特征名

def incomecla1(income_level,lit):  # 收入层次分类器
    for i in lit:
        if i > 100000:
            income_level['D'] = income_level.get('D') + 1
        elif i > 50000:
            income_level['C'] = income_level.get('C') + 1
        elif i >10000 :
            income_level['B'] = income_level.get('B') + 1
        elif i > 1 :
            income_level['A'] = income_level.get('A') + 1
    return income_level

def redit(dit,num):  # 重构字典
    if num == 0:
        return
    for i in dit:
        dit[i] = '{:.3f}%'.format(dit.get(i)/num*100)
    return dit

def incomecla2(i,x):  # 收入层次分类器改良版
    if i > 100000:
        x['D'] = x.get('D') + 1
    elif i > 50000:
        x['C'] = x.get('C') + 1
    elif i > 10000:
        x['B'] = x.get('B') + 1
    elif i > 1:
        x['A'] = x.get('A') + 1
    return x

income = table.col_values(-1)[1:]
Malenum = 0
Femalenum = 0
income_gender1 = {'A':0,'B':0,'C':0,'D':0}  # 收入层次分为ABCD四类，男
income_gender2 = {'A':0,'B':0,'C':0,'D':0}  # 收入层次分为ABCD四类，女
gender = table.col_values(0)[1:]  # 性别

def gendercla():  # 性别分类器
    global Malenum,Femalenum
    global income_gender1,income_gender2
    for i in range(len(gender)):
        if gender[i] == 1:
            Malenum += 1
            incomecla2(income[i],income_gender1)
        else:
            Femalenum += 1
            incomecla2(income[i], income_gender2)
    redit(income_gender1, Malenum)
    redit(income_gender2, Femalenum)
    print('收入阶层-性别')
    print(Malenum, Femalenum)  # 男 56 女42
    print(income_gender1)
    print(income_gender2)
    return

gendercla()


agenum1 = 0  # 0-29
agenum2 = 0  # 30-49
agenum3 = 0  # 50-
income_age1 = {'A':0,'B':0,'C':0,'D':0}
income_age2 = {'A':0,'B':0,'C':0,'D':0}
income_age3 = {'A':0,'B':0,'C':0,'D':0}
age = table.col_values(1)[1:]

def agecla():  # 年龄分类器
    global agenum1,agenum2,agenum3
    global income_age1, income_age2, income_age3
    for i in range(len(age)):
        if age[i] >= 50:
            agenum3 += 1
            income_age3 = incomecla2(income[i],income_age3)
        elif age[i]>= 30:
            agenum2 += 1
            income_age2 = incomecla2(income[i],income_age2)
        else:
            agenum1 += 1
            income_age1 = incomecla2(income[i],income_age1)
    redit(income_age1, agenum1)
    redit(income_age2, agenum2)
    redit(income_age3, agenum3)
    print('收入阶层-年龄')
    print(agenum1, agenum2, agenum3)
    print(income_age1, income_age2, income_age3)
    return

agecla()


workdays1 = 0  # 0-15
workdays2 = 0  # 16-30
income_work1 = {'A':0,'B':0,'C':0,'D':0}
income_work2 = {'A':0,'B':0,'C':0,'D':0}
workdays = table.col_values(9)[1:]

def workcla():  # 工作天数分类器
    global workdays1,workdays2
    global income_work1,income_work2
    for i in range(len(workdays)):
        if workdays[i] >= 16:
            workdays2 += 1
            income_work2 = incomecla2(income[i],income_work2)
        else:
            workdays1 += 1
            income_work1 = incomecla2(income[i],income_work1)
    redit(income_work1, workdays1)
    redit(income_work2, workdays2)
    print('收入阶层-工作天数')
    print(workdays1, workdays2)
    print(income_work1, income_work2)
    return

workcla()


workhours1 = 0  # 0-7
workhours2 = 0  # 8-
income_hours1 = {'A':0,'B':0,'C':0,'D':0}
income_hours2 = {'A':0,'B':0,'C':0,'D':0}
workhours = table.col_values(10)[1:]

def workhourscla():  # 日均工作时间分类器
    global workhours1,workhours2
    global income_hours1,income_hours2
    for i in range(len(workhours)):
        if workhours[i] > 7:
            workhours2 += 1
            income_hours2 = incomecla2(income[i],income_hours2)
        else:
            workhours1 += 1
            income_hours1 = incomecla2(income[i], income_hours1)
    redit(income_hours1, workhours1)
    redit(income_hours2, workhours2)
    print('收入阶层-日均工作时间')
    print(workhours1, workhours2)
    print(income_hours1, income_hours2)
    return

workhourscla()


Edu_level = table.col_values(2)[1:]
Edu_level_num = [0,0,0,0,0,0,0,0,0]
income_edu1 = {'A':0,'B':0,'C':0,'D':0}
income_edu2 = {'A':0,'B':0,'C':0,'D':0}
income_edu3 = {'A':0,'B':0,'C':0,'D':0}
income_edu4 = {'A':0,'B':0,'C':0,'D':0}
income_edu5 = {'A':0,'B':0,'C':0,'D':0}
income_edu6 = {'A':0,'B':0,'C':0,'D':0}
income_edu7 = {'A':0,'B':0,'C':0,'D':0}
income_edu8 = {'A':0,'B':0,'C':0,'D':0}
income_edu9 = {'A':0,'B':0,'C':0,'D':0}

def educal():  # 学历水平分类器
    global Edu_level_num
    global income_edu1,income_edu2,income_edu3,income_edu4,income_edu5,income_edu6,income_edu7,income_edu8,income_edu9
    for i in range(len(Edu_level)):
        if Edu_level[i] == 1:
            Edu_level_num[0] +=1
            income_edu1 = incomecla2(income[i],income_edu1)
        elif Edu_level[i] == 2:
            Edu_level_num[1] += 1
            income_edu2 = incomecla2(income[i],income_edu2)
        elif Edu_level[i] == 3:
            Edu_level_num[2] += 1
            income_edu3 = incomecla2(income[i],income_edu3)
        elif Edu_level[i] == 4:
            Edu_level_num[3] += 1
            income_edu4 = incomecla2(income[i],income_edu4)
        elif Edu_level[i] == 5:
            Edu_level_num[4] += 1
            income_edu5 = incomecla2(income[i],income_edu5)
        elif Edu_level[i] == 6:
            Edu_level_num[5] += 1
            income_edu6 = incomecla2(income[i],income_edu6)
        elif Edu_level[i] == 7:
            Edu_level_num[6] += 1
            income_edu7 = incomecla2(income[i],income_edu7)
        elif Edu_level[i] == 8:
            Edu_level_num[7] += 1
            income_edu8 = incomecla2(income[i],income_edu8)
        else:
            Edu_level_num[8] += 1
            income_edu9 = incomecla2(income[i],income_edu9)
    redit(income_edu1, Edu_level_num[0])
    redit(income_edu2, Edu_level_num[1])
    redit(income_edu3, Edu_level_num[2])
    redit(income_edu4, Edu_level_num[3])
    redit(income_edu5, Edu_level_num[4])
    redit(income_edu6, Edu_level_num[5])
    redit(income_edu7, Edu_level_num[6])
    redit(income_edu8, Edu_level_num[7])
    redit(income_edu9, Edu_level_num[8])
    print(('收入阶层-学历水平'))
    print(Edu_level_num)
    print(income_edu1)
    print(income_edu2)
    print(income_edu3)
    print(income_edu4)
    print(income_edu5)
    print(income_edu6)
    print(income_edu7)
    print(income_edu8)
    print(income_edu9)
    return

educal()


communist = table.col_values(3)[1:]
communist_num = [0,0]
communist1 = {'A':0,'B':0,'C':0,'D':0}
communist2 = {'A':0,'B':0,'C':0,'D':0}

def comcal():  # 共产党员分类器
    global communist_num
    global communist1,communist2
    for i in range(len(communist)):
        if communist[i] == 1:
            communist_num[0] += 1
            communist1 = incomecla2(income[i],communist1)
        else:
            communist_num[1] += 1
            communist2 = incomecla2(income[i], communist2)
    redit(communist1, communist_num[0])
    redit(communist2, communist_num[1])
    print('收入阶层-是否为共产党员')
    print(communist_num)
    print(communist1)
    print(communist2)
    return

comcal()

reg_res = table.col_values(4)[1:]
reg_res_num = [0,0,0]
reg_res1 = {'A':0,'B':0,'C':0,'D':0}
reg_res2 = {'A':0,'B':0,'C':0,'D':0}
reg_res3 = {'A':0,'B':0,'C':0,'D':0}

def reg_res_cal():  # 户口类型分类器
    global reg_res_num
    global reg_res1,reg_res2,reg_res3
    for i in range(len(reg_res)):
        if reg_res[i] == 1:
            reg_res_num[0] += 1
            reg_res1 = incomecla2(income[i],reg_res1)
        elif reg_res[i] == 2:
            reg_res_num[1] += 1
            reg_res2 = incomecla2(income[i], reg_res2)
        else:
            reg_res_num[2] += 1
            reg_res3 = incomecla2(income[i], reg_res3)
    redit(reg_res1, reg_res_num[0])
    redit(reg_res2, reg_res_num[1])
    redit(reg_res3, reg_res_num[2])
    print('收入阶层-户口类型')
    print(reg_res_num)
    print(reg_res1)
    print(reg_res2)
    print(reg_res3)
    return

reg_res_cal()

marriage = table.col_values(5)[1:]
marriage_num = [0,0]
marriage1 = {'A':0,'B':0,'C':0,'D':0}
marriage2 = {'A':0,'B':0,'C':0,'D':0}

def marriage_cal():  # 婚姻状况分类器
    global marriage
    global marriage1,marriage2
    for i in range(len(marriage)):
        if marriage[i] == 1:
            marriage_num[0] += 1
            marriage1 = incomecla2(income[i],marriage1)
        else:
            marriage_num[1] += 1
            marriage2 = incomecla2(income[i],marriage2)
    redit(marriage1,marriage_num[0])
    redit(marriage2,marriage_num[1])
    print('收入阶层-婚姻状况')
    print(marriage_num)
    print(marriage1)
    print(marriage2)
    return

marriage_cal()

healthy = table.col_values(6)[1:]
healthy_num = [0,0,0,0]
healthy1 = {'A':0,'B':0,'C':0,'D':0}
healthy2 = {'A':0,'B':0,'C':0,'D':0}
healthy3 = {'A':0,'B':0,'C':0,'D':0}
healthy4 = {'A':0,'B':0,'C':0,'D':0}

def healthy_cal():  # 身体状况分类器
    global healthy_num
    global healthy1,healthy2,healthy3,healthy4
    for i in range(len(healthy)):
        if healthy[i] == 1:
            healthy_num[0] += 1
            healthy1 = incomecla2(income[i], healthy1)
        elif healthy[i] == 2:
            healthy_num[1] += 1
            healthy2 = incomecla2(income[i], healthy2)
        elif healthy[i] == 3:
            healthy_num[2] += 1
            healthy3 = incomecla2(income[i], healthy3)
        else:
            healthy_num[3] += 1
            healthy4 = incomecla2(income[i], healthy4)
    redit(healthy1, healthy_num[0])
    redit(healthy2, healthy_num[1])
    redit(healthy3, healthy_num[2])
    redit(healthy4, healthy_num[3])
    print('收入阶层-身体状况')
    print(healthy_num)
    print(healthy1)
    print(healthy2)
    print(healthy3)
    print(healthy4)
    return

healthy_cal()


hangye = table.col_values(7)[1:]
hangye_num = [0]*24
hangye_cla = [{'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              {'A':0,'B':0,'C':0,'D':0},
              ]

def hangye_cal():  # 所在行业分类器
    global hangye_num,hangye_cla
    for i in range(len(hangye)):
        if hangye[i] == 1:
            hangye_num[0] += 1
            hangye_cla[0] = incomecla2(income[i],hangye_cla[0])
        elif hangye[i] == 2:
            hangye_num[1] += 1
            hangye_cla[1] = incomecla2(income[i],hangye_cla[1])
        elif hangye[i] == 3:
            hangye_num[2] += 1
            hangye_cla[2] = incomecla2(income[i],hangye_cla[2])
        elif hangye[i] == 4:
            hangye_num[3] += 1
            hangye_cla[3] = incomecla2(income[i],hangye_cla[3])
        elif hangye[i] == 5:
            hangye_num[4] += 1
            hangye_cla[4] = incomecla2(income[i],hangye_cla[4])
        elif hangye[i] == 6:
            hangye_num[5] += 1
            hangye_cla[5] = incomecla2(income[i],hangye_cla[5])
        elif hangye[i] == 7:
            hangye_num[6] += 1
            hangye_cla[6] = incomecla2(income[i],hangye_cla[6])
        elif hangye[i] == 8:
            hangye_num[7] += 1
            hangye_cla[7] = incomecla2(income[i],hangye_cla[7])
        elif hangye[i] == 9:
            hangye_num[8] += 1
            hangye_cla[8] = incomecla2(income[i],hangye_cla[8])
        elif hangye[i] == 10:
            hangye_num[9] += 1
            hangye_cla[9] = incomecla2(income[i],hangye_cla[9])
        elif hangye[i] == 11:
            hangye_num[10] += 1
            hangye_cla[10] = incomecla2(income[i],hangye_cla[10])
        elif hangye[i] == 12:
            hangye_num[11] += 1
            hangye_cla[11] = incomecla2(income[i],hangye_cla[11])
        elif hangye[i] == 13:
            hangye_num[12] += 1
            hangye_cla[12] = incomecla2(income[i],hangye_cla[12])
        elif hangye[i] == 14:
            hangye_num[13] += 1
            hangye_cla[13] = incomecla2(income[i],hangye_cla[13])
        elif hangye[i] == 15:
            hangye_num[14] += 1
            hangye_cla[14] = incomecla2(income[i],hangye_cla[14])
        elif hangye[i] == 16:
            hangye_num[15] += 1
            hangye_cla[15] = incomecla2(income[i],hangye_cla[15])
        elif hangye[i] == 17:
            hangye_num[16] += 1
            hangye_cla[16] = incomecla2(income[i],hangye_cla[16])
        elif hangye[i] == 18:
            hangye_num[17] += 1
            hangye_cla[17] = incomecla2(income[i],hangye_cla[17])
        elif hangye[i] == 19:
            hangye_num[18] += 1
            hangye_cla[18] = incomecla2(income[i],hangye_cla[18])
        elif hangye[i] == 20:
            hangye_num[19] += 1
            hangye_cla[19] = incomecla2(income[i],hangye_cla[19])
        elif hangye[i] == 21:
            hangye_num[20] += 1
            hangye_cla[20] = incomecla2(income[i],hangye_cla[20])
        elif hangye[i] == 22:
            hangye_num[21] += 1
            hangye_cla[21] = incomecla2(income[i],hangye_cla[21])
        elif hangye[i] == 23:
            hangye_num[22] += 1
            hangye_cla[22] = incomecla2(income[i],hangye_cla[22])
        else:
            hangye_num[23] += 1
            hangye_cla[23] = incomecla2(income[i],hangye_cla[23])
    print('收入阶层-所在行业')
    print(hangye_num)
    for i in range(len(hangye_num)):
        redit(hangye_cla[i],hangye_num[i])
        print(hangye_cla[i])
    return

hangye_cal()

industry = table.col_values(8)[1:]
industry_num = [0,0,0]
industry_cla = [{'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},]

def industry_cal():  # 所在产业分类器
    global industry_num,industry_cla
    for i in range(len(industry)):
        if industry[i] == 1:
            industry_num[0] += 1
            industry_cla[0] = incomecla2(income[i],industry_cla[0])
        elif industry[i] == 2:
            industry_num[1] += 1
            industry_cla[1] = incomecla2(income[i],industry_cla[1])
        else:
            industry_num[2] += 1
            industry_cla[2] = incomecla2(income[i], industry_cla[2])
    print('收入阶层-所在行业')
    print(industry_num)
    for i in range(len(industry_cla)):
        redit(industry_cla[i],industry_num[i])
        print(industry_cla[i])
    return

industry_cal()

job_tybe = table.col_values(-2)[1:]
job_tybe_num = [0]*10
job_tybe_cla = [{'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},
                {'A':0,'B':0,'C':0,'D':0},]

def job_tybe_cal():  # 工作类型分类器
    global job_tybe_num,job_tybe_cla
    for i in range(len(job_tybe)):
        if job_tybe[i] == 1:
            job_tybe_num[0] += 1
            job_tybe_cla[0] = incomecla2(income[i],job_tybe_cla[0])
        elif job_tybe[i] == 2:
            job_tybe_num[1] += 1
            job_tybe_cla[1] = incomecla2(income[i],job_tybe_cla[1])
        elif job_tybe[i] == 3:
            job_tybe_num[2] += 1
            job_tybe_cla[2] = incomecla2(income[i],job_tybe_cla[2])
        elif job_tybe[i] == 4:
            job_tybe_num[3] += 1
            job_tybe_cla[3] = incomecla2(income[i],job_tybe_cla[3])
        elif job_tybe[i] == 5:
            job_tybe_num[4] += 1
            job_tybe_cla[4] = incomecla2(income[i],job_tybe_cla[4])
        elif job_tybe[i] == 6:
            job_tybe_num[5] += 1
            job_tybe_cla[5] = incomecla2(income[i],job_tybe_cla[5])
        elif job_tybe[i] == 7:
            job_tybe_num[6] += 1
            job_tybe_cla[6] = incomecla2(income[i],job_tybe_cla[6])
        elif job_tybe[i] == 8:
            job_tybe_num[7] += 1
            job_tybe_cla[7] = incomecla2(income[i],job_tybe_cla[7])
        elif job_tybe[i] == 9:
            job_tybe_num[8] += 1
            job_tybe_cla[8] = incomecla2(income[i],job_tybe_cla[8])
        else:
            job_tybe_num[9] += 1
            job_tybe_cla[9] = incomecla2(income[i],job_tybe_cla[9])
    print('收入阶层-工作类型')
    print(job_tybe_num)
    for i in range(len(job_tybe_cla)):
        redit(job_tybe_cla[i],job_tybe_num[i])
        print(job_tybe_cla[i])
    return

job_tybe_cal()