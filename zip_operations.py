import shutil
import os

folder_name = input("Enter folder name: ")
archive_name = input("Enter archive file name (without extension): ")

if os.path.exists(folder_name):
    shutil.make_archive(archive_name, 'zip', folder_name)
    print("Folder zipped successfully")
else:
    print("Error: Folder does not exist!")
