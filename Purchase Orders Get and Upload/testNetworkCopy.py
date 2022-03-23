import shutil


source_path = r"\\mynetworkshare"
dest_path = r"C:\TEMP"
file_name = "\\myfile.txt"

shutil.copyfile(source_path + file_name, dest_path + file_name)
