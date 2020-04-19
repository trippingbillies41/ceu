import os, sys
from pdf2image import convert_from_path
from PIL import Image
import csv
import pandas as pd
import shutil
import os.path
import pickle
import numpy as np
import pytesseract

# Print Instructions

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
error_destination = os.path.join(in_path, "ERRORS")
# Estalish Employee Lists
emplist = []
emplist_wval = []
values = []

# Table Directory
if dir_pkl_s == None:
	dir_table = input('Table Directory: ')
else:
	dir_table = (dir_pkl_s[2])

# Loading Names From Excel Table
df = pd.read_csv(dir_table)
emplist = df['Name'].tolist()

# Certificate Directory
if dir_pkl_s == None:
	dir_cert = input('Employee Certificate Parent Directory: ')
else:
	dir_cert = (dir_pkl_s[3])
# Check Names vs Folders
subfolders = os.listdir(dir_cert)
os.chdir(dir_cert)
for name in emplist:
	if name in subfolders:
		pass
	else:
		os.mkdir(name)
		print("Folder", name, "Created in Employee Certificate Directory.")

os.chdir(in_path)

# Creating Values List
val_input = input('Load Previous Values? y or n: ')
if val_input == 'y':
	values = list(np.load('values.npy'))
else:
	values = ['startingplaceholder']
val_add_input = input('Add New Values? y or n: ')
if val_add_input == "y":
	val = ""
	print('Values Must Start With "@".')
	print('Mandatories (ex: @1@M1) Must Be Added First.')
	print('Type "done" to Finish.')
	while val != "done":
		val = input("Add Values With '@' as a Delimiter: ")
		if val in values:
			print('This Value Already Exists')
		else:
			values.append(val)
	values.remove('done')
	if 'startingplaceholder' in values:
		values.remove('startingplaceholder')
	else:
		pass
else:
	pass

# Names with Values
emplist_wval = [(e+v) for e in emplist for v in values]

# Define Functions

# Create JPEG
def certificate_counter(x):
	out_jpeg = x[:-3] + "jpeg"
	outf = convert_from_path(x)
	for out in outf:
		out.save(out_jpeg, 'JPEG')
	# Read TXT from JPEG
	text = pytesseract.image_to_string(Image.open(out_jpeg))
	text = text.replace('-\n', '')
	# Create TXT File
	outfiletext = (x[:-4] + ".txt")
	f = open(outfiletext, "w")
	f.write(text)
	f.close()
	t = open(outfiletext, 'r')
	readtext = t.read()
	for iteration in set(emplist_wval):
		try:
			if iteration in readtext:
				certificatename = iteration
				delim_count = certificatename.count('@')
				if delim_count == 1:
					cert_name, cert_value = certificatename.split("@")
					print(cert_name, cert_value)
					# Rename PDF
					for pdf in os.listdir(in_path):
						arg_1 = pdf[-3:] == "pdf"
						arg_2 = pdf[:-4] == (outfiletext[:-4])
						if arg_1 and arg_2 == True:
							root_name = outfiletext[:-4]
							old_name = outfiletext[:-4] + ".pdf"
							new_name = cert_name + "!" + root_name + ".pdf"
							pdf = os.rename(old_name, new_name)
			 		# Open Table
					df2 = pd.read_csv(dir_table, index_col=[0])
					# Cells
					cell_v = df2.loc[cert_name, "CEU"]
					cell_m1 = df2.loc[cert_name, "M1"]
					cell_m2 = df2.loc[cert_name, "M2"]
					# Adding CEU Value
					df2.loc[cert_name, "CEU"] = (cell_v + float(cert_value))
					# Save Table
					df2.to_csv(dir_table)
					# Move Certificate to Employee Files
					destination = os.path.join(dir_cert, cert_name)
					move_pdf = cert_name + "!" + x
					shutil.move(move_pdf, destination)
					# Clear Variables
					cert_name = None
					cert_value = None
				elif delim_count == 2:
					cert_name, cert_value, cert_m = certificatename.split("@")	
					print(cert_name, cert_value, cert_m)
					# Rename PDF
					for pdf in os.listdir(in_path):
						arg_1 = pdf[-3:] == "pdf"
						arg_2 = pdf[:-4] == (outfiletext[:-4])
						if arg_1 and arg_2 == True:
							root_name = outfiletext[:-4]
							old_name = outfiletext[:-4] + ".pdf"
							new_name = cert_name + "!" + root_name + ".pdf"
							pdf = os.rename(old_name, new_name)
			 		# Open Table
					df2 = pd.read_csv(dir_table, index_col=[0])
					# Cells
					cell_v = df2.loc[cert_name, "CEU"]
					cell_m1 = df2.loc[cert_name, "M1"]
					cell_m2 = df2.loc[cert_name, "M2"]
					# Adding CEU Value
					df2.loc[cert_name, "CEU"] = (cell_v + float(cert_value))
					# Assigning Mandatories
					try:
						if cert_m == 'M1':
							df2.loc[cert_name, "M1"] = 'X'
						elif cert_m == 'M2':
							df2.loc[cert_name, "M2"] = 'X'
					except NameError:
						pass
					# Save Table
					df2.to_csv(dir_table)
					# Move Certificate to Employee Files
					destination = os.path.join(dir_cert, cert_name)
					move_pdf = cert_name + "!" + x
					shutil.move(move_pdf, destination)
					# Clear Variables
					cert_name = None
					cert_value = None
					cert_m = None
			else:
				pass
		except:
			print('ERROR: Certificate Unreadable.')
			error_jpeg = outfiletext[:-4] + ".jpeg"
			os.remove(error_jpeg)
			shutil.move(x, error_destination)
			break

			
input_files = os.listdir(in_path)
for x in input_files:
	if x[-3:] == "pdf":
		certificate_counter(x)
		# Delete JPEG and TXT
		remove_jpeg = x[:-3] + "jpeg"
		remove_txt = x[:-3] + "txt"
		os.remove(remove_jpeg)
		os.remove(remove_txt)
	else:
		pass



