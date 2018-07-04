import random
import pprint
import time
import xlrd
from tkinter import *

master = Tk()
Label(master, text="Number of Students: ").grid(row=0)
Label(master, text="Number of Shifts: ").grid(row=1)
e1 = Entry(master)
e2 = Entry(master)
e1.grid(row=0, column=1)
e2.grid(row=1, column=1)

def set():
	#numberOfStudents = int(input("Enter the total nuumber of employees"))
	#numberOfShifts = int(input("Enter the total number of shifts"))
	if ( e1.get() == "5" ):
		numberOfStudents = 5		
	if ( e2.get() == "8" ):
		numberOfShifts = 8
	master.destroy()
	estudents = {}
	for i in range ( 1, numberOfStudents+1 ):
		estudents[i] = [[],[],[]]
	students = {}
	for i in range ( 1, numberOfStudents+1 ):
		students[i] = [[],[],[]]
	file_location = "./data.xlsx"
	workbook = xlrd.open_workbook(file_location)
	sheet = workbook.sheet_by_index(0)
	for rows in range(1, sheet.nrows ):
		for cols in range(2, sheet.ncols ):
			estudents[rows][cols-2] = (sheet.cell_value(rows,cols)).split(",")
	names = {}
	for rows in range(1, sheet.nrows):
		names[rows] = sheet.cell_value(rows,0)
	for copy in range(1, numberOfStudents + 1):
		for lists in range(3):
			for items in range( len( estudents[copy][lists] ) ):
				if ( estudents[copy][lists][items] != "" ):
					if ( estudents[copy][lists][items][-1] ) == "1":
						temp = estudents[copy][lists][items][0]
						temp = int (temp)
						temp = (temp*2)-1
					if ( estudents[copy][lists][items][-1] ) == "2":
						temp = estudents[copy][lists][items][0]
						temp = int (temp)
						temp = (temp*2)
					students[copy][lists].append(temp)
	shifts = {}
	for i in range ( 1, numberOfShifts+1 ):
		shifts[i] = []
	track = {}
	for i in range ( 1, numberOfStudents+1 ):
		track[i] = 0
	schedule = {}
	for i in range ( 1, numberOfStudents+1 ):
		schedule[i] = []
	for day in range(1, numberOfShifts+1):
		checked = []
		counter = 0 
		while(counter < 2 and len(checked) < numberOfStudents):
			randomStudent = random.randint(1, numberOfStudents)
			if randomStudent in checked:
				continue
			checked.append(randomStudent)
			if day in students[randomStudent][0]:
				shifts[day].append(randomStudent)
				counter += 1
				track[randomStudent] += 1
				schedule[randomStudent].append(day)
	for day in range(1,numberOfShifts+1):
		if( len ( shifts[day] ) == 2 ):
			continue
		checked = []
		counter = 0 
		while(counter < 2 and len(checked) < numberOfStudents):
			randomStudent = random.randint(1, numberOfStudents)
			if randomStudent in checked:
				continue
			checked.append(randomStudent)
			if day in students[randomStudent][1]:
				#shifts[day] = [randomStudent]
				shifts[day].append(randomStudent)
				counter += 1
				track[randomStudent] += 1
				schedule[randomStudent].append(day)
	# now let's fill the empty / incomplete days
	tries = 0
	for day in range(1, numberOfShifts+1):
		if( len ( shifts[day] ) == 2 ):
			continue
		while( ( len ( shifts[day] ) < 2 ) and tries < len(students) ):
			randomStudent = random.randint(1, numberOfStudents)	
			if day not in students[randomStudent][2]:
				if shifts[day][0] == randomStudent:
					tries += 1
					continue
				shifts[day].append(randomStudent)
				track[randomStudent] += 1
				break
	# equalling shifts
	idealShifts = (len(shifts)*2) // len(students)
	# print("Shifts per person: ")
	# print(idealShifts)
	for employee in track:
		while track[employee] < (idealShifts):
			# print ( str(employee) + " doesn't have ideal # of shifts" )
			maxStudent = max(track, key=track.get) 
			if track[maxStudent] == idealShifts:
				break
			# print(str(maxStudent) + " has the maximum number of shifts") 
			# print("shifts of max student")
			# print( schedule[maxStudent] )
			flag = True
			for days in schedule[maxStudent]:
				if flag == False:
					continue
				if days not in schedule[employee]:
					if days not in students[employee][2]:
						#remove this day from max student
						schedule[maxStudent].remove(days)
						track[maxStudent] -= 1
						#add this day to current employee
						schedule[employee].append(days)
						track[employee] += 1
						#removing maxstudent from that shift
						shifts[days].remove(maxStudent)
						#adding current employee to that shift
						shifts[days].append(employee)
						flag = False
	# yes print("======= F I N A L =======")	
	# yes print("The schedule looks like: ")
	# yes print(shifts)	
	# yes print("Individul shift tracking: ")
	# yes print(track)
	for employee in track:
		while track[employee] > (idealShifts+1):
			# print ( str(employee) + " has more than ideal # of shifts" )
			minStudent = min(track, key=track.get) 
			#if track[maxStudent] == idealShifts:
			#	break
			# print(str(minStudent) + " has the minimum number of shifts") 
			# print("shifts of min student")
			# print( schedule[minStudent] )
			flag = True
			for days in schedule[employee]:
				if flag == False:
					continue
				if days not in schedule[minStudent]:
					if days not in students[minStudent][2]:
						#remove this day from employee
						schedule[employee].remove(days)
						track[employee] -= 1
						#add this day to min student
						schedule[minStudent].append(days)
						track[minStudent] += 1
						#removing employee from that shift
						shifts[days].remove(employee)
						#adding min student to that shift
						shifts[days].append(minStudent)
						flag = False
	#printing shifts
	'''
	for key in shifts.keys():
		date = (key + 1) //2
		if ( len (shifts[key] ) == 2 ):
			# print( str(shifts[key][0]) + " and " + str(shifts[key][1]) + " work on " + str(date) )
		elif ( len (shifts[key] ) == 1 ):
			# yes print( str(shifts[key][0]) + " works on " + str(date) )
		else:
			# yes print("noone works on " + str(date) )
	#print("\n".join("{}\t{}".format(k, v) for k, v in shifts.items()))
	'''
	print("*******************************************")
	print(" F I N A L  S C H E D U L E ")
	print("*******************************************")
	for shift in range ( 1, len(shifts) + 1 ):
		#print(str( shift ) + ": " )
		if shift%2 != 0:
			date = (shift+1)/2
			oddeven = "odd"
		else:
			date = shift/2
			oddeven = "even"
		if oddeven == "odd":
			print( "Day " + str( date )[0] + " shift 1" )
		else:
			print( "Day " + str( date )[0] + " shift 2" )
		temp = shifts[shift] 
		#print(temp)
		if len(temp) == 2:
			print(  names[ temp [0] ]  + ", " +names [ temp[1] ] )
		else:
			print( names [ temp [0] ]  )
		print()
	print()
	print("Shifts for each person: ")
	print()
	for id in track:
		print( names[id] + " has " + str( track[id] ) + " shifts" )
Button(master, text='Go', command=set).grid(row=3, column=0, sticky=W, pady=4)
mainloop( )
