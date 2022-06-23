import os

def main():

    global parent_os
    parent_os = os.path.dirname(os.path.realpath(__file__))
    open_file1 = parent_os + '/excelReader.py'
    open_file2 = parent_os + '/outputIcs.py'
    os.system("python " + open_file1)
    os.system("python " + open_file2)

if __name__ == "__main__":
    main()