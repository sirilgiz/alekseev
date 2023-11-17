# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

def create_null_files ():
    f = open('filelist.txt','r')
    for filename in f:
        sFilename = "art\\" + filename.rstrip()
        f2 = open(sFilename,'w')
        f2.close

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    create_null_files()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/



