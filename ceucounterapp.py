import os, sys
# import pypdf2
import csv
import pandas as pd
import shutil
import os.path
import pickle
import numpy as np
#WINDOWS
# Establish Directories
dir_input = input('Load Previous Directories? y or n: ')
if dir_input == 'y':
	dir_pkl = open("dir_pkl.pickle", "rb")
	dir_pkl_s = pickle.load(dir_pkl)
else:
	dir_pkl_s = None

if dir_pkl_s == None:
	new_dir_input = input('Certificate Input Directory: ')
else:
	new_dir_input = (dir_pkl_s[1])
in_path = (new_dir_input)

# Change Working Directory
os.chdir(in_path)

# Estalish Employee Lists
emplist = []
emplist_wval = []
emptemplist = []
values = []

# Table Directory
if dir_pkl_s == None:
	dir_table = input('Table Directory: ')
else:
	dir_table = (dir_pkl_s[2])

# Loading Names From Excel Table
df = pd.read_csv(dir_table)
print(df)
emplist = df['Name'].tolist()

# Creating Values List
val_input = input('Load Previous Values? y or n: ')
if val_input == 'y':
	values = list(np.load('values.npy'))
else:
	values = ['startingplaceholder']
val_add_input = input('Add New Values? y or n: ')
if val_add_input == "y":
	val = ""
	while val != "done":
		val = input("Add Values With '@' as a Delimiter: ")
		values.append(val)
	values.remove('done')
	if 'startingplaceholder' in values:
		values.remove('startingplaceholder')
	else:
		pass
else:
	pass

print(emplist)
print(values)
# Names with Values
emplist_wval = [(e+v) for e in emplist for v in values]

print(emplist_wval)

# # Create JPEG Files
# for infile in os.listdir(in_path):
#     if infile[-3:] == "tif":
# 	    outfile = infile[:-3] + "jpeg"
# 	    im = tifffile.imread(infile)
# 	    out = im.convert("RGB")
# 	    out.save(outfile, "JPEG", quality=90)

# # Create TXT Files
# for cert_jpeg in os.listdir(in_path):
# 	if cert_jpeg[-4:] == "jpeg":


# 		text = pytesseract.image_to_string(Image.open(cert_jpeg))
# 		text = text.replace('-\n', '')

# 		# Create Output Text File
# 		outfiletext = (cert_jpeg + ".txt")
# 		f = open(outfiletext, "w") 
# 		f.write(text) 
# 		f.close()

# # Count Certificates

# for txt_file in os.listdir(in_path):
# 	if txt_file[-3:] == "txt":
# 		for name in emplist_wval:
# 			with open(txt_file) as f:
# 				if name in f.read():
# 					certificatename = name
# 					break
# 				else:
# 					certificatename = "Not_Found"

# 		delim_count = certificatename.count('@')
# 		if delim_count == 1:
# 			cert_name, cert_value = certificatename.split("@")
# 		elif delim_count == 2:
# 			cert_name, cert_value, cert_m = certificatename.split("@")		
		
# 		df = pd.read_csv(dir_table, index_col=[0])
# 		for tif in os.listdir(in_path):
# 			arg_1 = tif[-3:] == "tif"
# 			arg_2 = tif[:-4] == (txt_file[:-9])
# 			if arg_1 and arg_2 == True:
# 				root_name = txt_file[:-9]
# 				old_name = txt_file[:-9] + ".tif"
# 				new_name = cert_name + "|" + root_name + ".tif"
# 				tif = os.rename(old_name, new_name)

# 		# Cells - change cert_name back to x
# 		cell_v = df.loc[cert_name, "CEU"]
# 		cell_m1 = df.loc[cert_name, "M1"]
# 		cell_m2 = df.loc[cert_name, "M2"]
# 		# Adding CEU Value
# 		df.loc[cert_name, "CEU"] = (cell_v + float(cert_value))
# 		# Assigning Mandatories
# 		try:
# 			if cert_m == 'M1':
# 				df.loc[cert_name, "M1"] = 'X'
# 			elif cert_m == 'M2':
# 				df.loc[cert_name, "M2"] = 'X'
# 		except NameError:
# 			pass
# # Clear Mandatory
# 		cert_m = None
# 		df.to_csv(table_dir)

# Certificate Directory
if dir_pkl_s == None:
	dir_cert = input('Employee Certificate Parent Directory')
else:
	dir_cert = (dir_pkl_s[3])

# # Move Certificate to Employee Files
# for move_file in os.listdir(in_path):
# 	if move_file[-3:] == "tif":
# 		move_file_name, garbage = (move_file.split("|"))
# 		desination = os.path.join(input_3, move_file_name)
# 		shutil.move(move_file, desination)

# # Delete TXT Files
# for remove_file in os.listdir(in_path):
# 	if remove_file[-3:] == "txt":
# 		os.remove(remove_file)

# # Delete JPEG Files
# for remove_file in os.listdir(in_path):
# 	if remove_file[-4:] == "jpeg":
# 		os.remove(remove_file)

# Create Save Files
save_question = input('Save Directories & Values? y or n: ')
if save_question == 'y':
	# Directories
	save_dict = {1:new_dir_input, 2:dir_table, 3:dir_cert}
	pickle_out = open("dir_pkl.pickle", "wb")
	pickle.dump(save_dict, pickle_out)
	pickle_out.close()
	# Names
	save_dict = emplist
	pickle_out = open("emp_pkl.pickle", "wb")
	pickle.dump(save_dict, pickle_out)
	pickle_out.close()
	# Values
	save_values = np.array(values)
	np.save('values', save_values)
else:
	pass