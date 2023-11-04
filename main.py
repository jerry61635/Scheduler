from datetime import datetime
import pandas as pd
import csv
import sys
import random
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.xlsx")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

def submit():
	# get current datetime
	dt = datetime.now()
	year = dt.year      #取得年份
	month = dt.month	#取得月份
	
	month_list = [31,28,31,30,31,30,31,31,30,31,30,31]
	day_name = ['一','二','三','四','五','六','日']
	staff_ID = []
	
	scheduler = []
	
	selected_file = file_path_entry.get()
	value = entry_value.get()
	
	xlsx_file = selected_file
	output_path = selected_file[:-1*len(selected_file.split('/')[-1])-1]
	print(output_path)

	read_file = pd.read_excel(xlsx_file)
	read_file.to_csv('convert.csv', index = None, header=True)

	user_input_date = value
	if user_input_date == '': user_input_date = str(year) + '/' + str(month+1)

	year = user_input_date.split('/')[0]
	month = user_input_date.split('/')[1]
	year = int(year)
	month = int(month)
	num_days = month_list[month-1] if year % 4 != 0 and month != 2 else 29

	print('+----------------------------------------------+')
	print(f"正在進行 {user_input_date} 的排班程序")
	print('+----------------------------------------------+')

	# 將轉檔後內容讀入
	with open('convert.csv', 'r') as file:
		reader = csv.reader(file)
		for line in reader:
			scheduler.append(line)

	# 確認誰要排班
	for i in range(len(scheduler[2])):
		if scheduler[2][i] == '隊員':
			staff_ID.append(i)

	#找誰公休
	chk={}
	for i in staff_ID: chk[i] = 0
	for i in range(5,num_days+5):
		for j in staff_ID:
			if scheduler[i][j] == '':
				chk[j] += 1

	# 刪除公休人員
	for i in range(len(staff_ID)-1,-1,-1):
		if chk[staff_ID[i]] == 0:
			staff_ID.pop(i)

	# 開始排班
	done = False
	threshold = 100
	average = num_days // len(staff_ID)
	quota = 1 if num_days % len(staff_ID) != 0 else 0
	while not done:
		scheduler = []
		with open('convert.csv', 'r') as file:
			reader = csv.reader(file)
			for line in reader:
				scheduler.append(line)
		retry = 0
		average_achive = True
		today, prev_day = {}, {}
		A_cnt, B_cnt, C_cnt, D_cnt, sum_cnt= {},{},{},{},{}
		for i in staff_ID: A_cnt[i] = 0; B_cnt[i] = 0; C_cnt[i] = 0; D_cnt[i] = 0; prev_day[i] = '0'; today[i] = '0'; sum_cnt[i] = 0

		for i in range(5, num_days+5):
			if retry >= threshold: break
			for staff in staff_ID: prev_day[staff] = '0'; today[staff] = '0'
			valid = False # 檢查生成的結果有沒有符合條件
			# 檢查這天誰能排班
			today_aval = []
			for staff in staff_ID:
				# 能排班的定義為0
				if scheduler[i][staff] != '': today[staff] = '1'
				else: today_aval.append(staff)
				if i > 5:
					if scheduler[i-1][staff] == '': prev_day[staff] = '0'
					elif scheduler[i-1][staff] == 'A': prev_day[staff] = 'A'
					elif scheduler[i-1][staff] == 'B': prev_day[staff] = 'B'
					elif scheduler[i-1][staff] == 'C': prev_day[staff] = 'C'
					elif scheduler[i-1][staff] == 'D': prev_day[staff] = 'D'
					elif scheduler[i-1][staff] != '': prev_day[staff] = '1'
				sum_cnt[staff] = A_cnt[staff] + B_cnt[staff] + C_cnt[staff] + D_cnt[staff]
			
			if i-4 > 20:
				print(f'Day{i-4}')
				print(f'yesterday: {prev_day}')
				print(f'today: {today}')
				print(f'summary: {sum_cnt}')
			while not valid:
				if retry >= threshold: break
				average_achive = True
				if len(today_aval) >= 4: pick = random.sample(today_aval, 4)
				else: retry = 10000; break
				
				if prev_day[pick[0]] != 'A' and prev_day[pick[0]] != 'B' and prev_day[pick[0]] != 'C' and today[pick[0]] != '1':
					if prev_day[pick[1]] != 'A' and prev_day[pick[1]] != 'B' and prev_day[pick[1]] != 'C' and today[pick[1]] != '1':
						if prev_day[pick[2]] != 'A' and prev_day[pick[2]] != 'B' and prev_day[pick[2]] != 'C' and today[pick[2]] != '1':
							if today[pick[3]] == '0': 
								if A_cnt[pick[0]] < average and B_cnt[pick[1]] < average and C_cnt[pick[2]] < average and D_cnt[pick[3]] < average:
									valid = True
								for staff in staff_ID:
									 if sum_cnt[staff] < average*4: average_achive = False
								if average_achive:
									if (sum_cnt[pick[0]] < average*4+quota and 
										sum_cnt[pick[1]] < average*4+quota and 
										sum_cnt[pick[2]] < average*4+quota and
										sum_cnt[pick[3]] < average*4+quota ):
										valid = True
				if not valid: retry += 1; continue
				retry = 0
				A_cnt[pick[0]] += 1
				B_cnt[pick[1]] += 1
				C_cnt[pick[2]] += 1
				D_cnt[pick[3]] += 1
				scheduler[i][pick[0]] = 'A'
				scheduler[i][pick[1]] = 'B'
				scheduler[i][pick[2]] = 'C'
				scheduler[i][pick[3]] = 'D'
				#print(pick)
				if i == num_days+4: done = True

	for staff in staff_ID:
		scheduler[37][staff] = A_cnt[staff]
		scheduler[38][staff] = B_cnt[staff]
		scheduler[39][staff] = C_cnt[staff]
		scheduler[40][staff] = D_cnt[staff]

	with open('scheduler.csv', 'w',encoding='utf-8' ,newline='') as out:
		writer = csv.writer(out)
		for line in scheduler:
			writer.writerow(line)

	read_file_product = pd.read_csv(f'scheduler.csv')
	read_file_product.to_excel(f'{output_path}/output.xlsx', index = None, header=True)
	messagebox.showinfo(f"排班完成!", f"排班完成!\n已將結果存在{output_path}/output.xlsx")

today = datetime.now().strftime("%Y/%m/%d")

root = tk.Tk()
root.title("排班工具")

file_path_label = tk.Label(root, text="選擇文件:")
file_path_label.pack()
file_path_entry = tk.Entry(root, width=40)
file_path_entry.pack()
browse_button = tk.Button(root, text="瀏覽", command=open_file_dialog)
browse_button.pack()

value_label = tk.Label(root, text=f"今天日期: {today}\n請輸入欲排班的月份yyyy/mm\n(若留空預設為當前下個月):")
value_label.pack()
entry_value = tk.Entry(root, width=20)
entry_value.pack()

submit_button = tk.Button(root, text="提交", command=submit)
submit_button.pack()

root.mainloop()
