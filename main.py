import pandas as pd
f = pd.read_excel('grade.xlsx')
name = f['课程名'].copy().values
score = f['分数'].copy().values
credit = f['学分'].copy().values
change = []
for i in range(0,70):
    score[i] = float(score[i])
    credit[i] = float(credit[i])
print(score)
print(credit)

sum_score = 0
sum_score_4 = 0
sum_credit = 0
now_gpa = 0
now_gpa_4 = 0
maxdecline_index = 0
maxdecline_num = 0
now_g = []
now_g_4 = []
to_4 = []
for i in range(0, 70):
    if(score[i] >= 85):
        to_4.append(4.0)
    elif(score[i] >= 75):
        to_4.append(3.0)
    elif(score[i] >= 60):
        to_4.append(2.0)
    else:
        to_4.append(0.0)




# 计算平均绩点
for i in range(0, 70):
    last_sum_score = sum_score  # 上一次的总分
    last_gpa = now_gpa # 上一次的平均绩点
    sum_score = sum_score + score[i] * credit[i] # 总分
    now_gpa = sum_score / (sum_credit + credit[i]) # 平均绩点
    sum_score_4 = sum_score_4 + to_4[i] * credit[i]
    now_gpa_4 = sum_score_4 / (sum_credit + credit[i])
    now_g.append(now_gpa)
    now_g_4.append(now_gpa_4)
    sum_credit = sum_credit + credit[i] # 总学分
    change.append(now_gpa - last_gpa) # 绩点变化
    if(maxdecline_num < last_gpa - now_gpa):
        maxdecline_num = last_gpa - now_gpa
        maxdecline_index = i

print('平均绩点为：', now_gpa)
print('平均绩点（4分制）为：', now_gpa_4)
print('最大下降绩点为：', maxdecline_num)
print('最大下降绩点的科目为：', name[maxdecline_index])

f['绩点变化'] = change
f['实时GPA'] = now_g
f['4分绩点'] = to_4
f['实时GPA（4分制）'] = now_g_4


f.to_excel('grade_new.xlsx')