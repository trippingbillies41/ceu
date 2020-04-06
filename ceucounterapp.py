import os, sys
from PIL import Image
import pytesseract
import csv
import pandas as pd
import shutil
import os.path
import pickle


start_input = input('Load Previous Directories? y or n')
if start_input == 'y':
	pickle_in = open("dict.pickle", "rb")
	save_dict = pickle.load(pickle_in)
else:
	save_dict = None

# INPUT_1
if save_dict == None:
	input_1 = input('Certificate Input Directory')
else:
	input_1 = (save_dict[1])
in_path = (input_1)

# Change Working Directory
os.chdir(in_path)

# Employee List
emplist = ['Merwede', 'Perez', 'Schmeck', 'Andrich', 'Jacob']
emplist_wval = ['Merwede@1', 'Merwede@1@2', 'Perez@1@M1', 'Perez@0.5']

# Create JPEG Files
for infile in os.listdir(in_path):
    if infile[-3:] == "tif":
	    outfile = infile[:-3] + "jpeg"
	    im = Image.open(infile)
	    out = im.convert("RGB")
	    out.save(outfile, "JPEG", quality=90)

# Create TXT Files
for cert_jpeg in os.listdir(in_path):
	if cert_jpeg[-4:] == "jpeg":


		text = pytesseract.image_to_string(Image.open(cert_jpeg))
		text = text.replace('-\n', '')

		# Create Output Text File
		outfiletext = (cert_jpeg + ".txt")
		f = open(outfiletext, "w") 
		f.write(text) 
		f.close()

# INPUT_2
if save_dict == None:
	input_2 = input('Table Directory')
else:
	input_2 = (save_dict[2])

# Count Certificates

for txt_file in os.listdir(in_path):
	if txt_file[-3:] == "txt":
		for name in emplist_wval:
			with open(txt_file) as f:
				if name in f.read():
					certificatename = name
					break
				else:
					certificatename = "Not_Found"

		delim_count = certificatename.count('@')
		if delim_count == 1:
			cert_name, cert_value = certificatename.split("@")
		elif delim_count == 2:
			cert_name, cert_value, cert_m = certificatename.split("@")		
		
		table_dir = (input_2)
		df = pd.read_csv(table_dir, index_col=[0])
		for tif in os.listdir(in_path):
			arg_1 = tif[-3:] == "tif"
			arg_2 = tif[:-4] == (txt_file[:-9])
			if arg_1 and arg_2 == True:
				root_name = txt_file[:-9]
				old_name = txt_file[:-9] + ".tif"
				new_name = cert_name + "|" + root_name + ".tif"
				tif = os.rename(old_name, new_name)

		# Cells - change cert_name back to x
		cell_v = df.loc[cert_name, "CEU"]
		cell_m1 = df.loc[cert_name, "M1"]
		cell_m2 = df.loc[cert_name, "M2"]
		# Adding CEU Value
		df.loc[cert_name, "CEU"] = (cell_v + float(cert_value))
		# Assigning Mandatories
		try:
			if cert_m == 'M1':
				df.loc[cert_name, "M1"] = 'X'
			elif cert_m == 'M2':
				df.loc[cert_name, "M2"] = 'X'
		except NameError:
			pass
# Clear Mandatory
		cert_m = None
		df.to_csv(table_dir)

# INPUT_3
if save_dict == None:
	input_3 = input('Employee Certificate Parent Directory')
else:
	input_3 = (save_dict[3])

# Move Certificate to Employee Files
for move_file in os.listdir(in_path):
	if move_file[-3:] == "tif":
		move_file_name, garbage = (move_file.split("|"))
		desination = os.path.join(input_3, move_file_name)
		shutil.move(move_file, desination)

# Delete TXT Files
for remove_file in os.listdir(in_path):
	if remove_file[-3:] == "txt":
		os.remove(remove_file)

# Delete JPEG Files
for remove_file in os.listdir(in_path):
	if remove_file[-4:] == "jpeg":
		os.remove(remove_file)

# Create Save File
save_dict = {1:input_1, 2:input_2, 3:input_3}
pickle_out = open("dict.pickle", "wb")
pickle.dump(save_dict, pickle_out)
pickle_out.close()