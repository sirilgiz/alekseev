import os
 
directory = 'd:\\Downloads\\slidoteka.ru\\all_files'
# target_directory=directory+'\'
for filename in os.listdir(directory):
    if os.path.isdir(filename):
        continue
    filename_clear = filename.split('(')[0] + '.' + filename.split(').')[1]
    #target_directory=directory+'\' + filename.split('(')[1].split(')')[0]
    target_directory = filename.split('@')[0].split('(')[1]
    target_directory=directory + '\\' + target_directory

    if not os.path.exists(target_directory):
        os.makedirs(target_directory)
    #print(directory+'\\'+filename+target_directory+'\\'+filename_clear)
    os.rename(directory + '\\'+filename,target_directory + '\\'+filename_clear)
