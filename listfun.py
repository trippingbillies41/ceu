emplist = []
emplist_wval = []
values = []

# Creating Names List
emps = ""
while emps != "done":
	emps = input("Add Employee Names, Type 'done' to Finish: ")
	emplist.append(emps)
emplist.remove('done')
print(emplist)

# Creating Values List
val = ""
while val != "done":
	val = input("Add Values With '@' as a Delimiter: ")
	values.append(val)
values.remove('done')

# Names with Values
emplist_wval = [(e+v) for e in emplist for v in values]
print(emplist_wval)