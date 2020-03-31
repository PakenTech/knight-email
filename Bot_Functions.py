import csv
import os
import shutil
import calendar
import datetime as date

import easygui as gui


def update_contacts(contact_import: str, output_file: str):
    # Open passed in contact file (reader) and output file (text to hold email string.)
    # Loop through contact file with emails in 5th column and extract to the output string.
    with open(contact_import) as csvfile:
        with open(output_file, 'w+') as email_txt:
            readCSV = csv.reader(csvfile, delimiter=",")
            email_list = ""
            loop = 0
            for row in readCSV:
                print(row)
                email = row[14]
                if loop == 0:
                    print("Header")
                elif loop == 1:
                    email_list = email
                elif email != "":
                    email_list = email_list + "; " + email
                print(email_list)
                loop += 1
            print("Writing file")
            email_txt.write(email_list)


def month_folder_path(month: str, parent_folder: str, year: str) -> (str, int):
    if month == "01" or month == "Jan" or month == "January":
        folder = year + " 01 - January"
        month_num = 1
    elif month == "02" or month == "Feb" or month == "February":
        folder = year + " 02 - February"
        month_num = 2
    elif month == "03" or month == "Mar" or month == "March":
        folder = year + " 03 - March"
        month_num = 3
    elif month == "04" or month == "Apr" or month == "April":
        folder = year + " 04 - April"
        month_num = 4
    elif month == "05" or month == "May":
        folder = year + " 05 - May"
        month_num = 5
    elif month == "06" or month == "Jun" or month == "June":
        folder = year + " 06 - June"
        month_num = 6
    elif month == "07" or month == "Jul" or month == "July":
        folder = year + " 07 - July"
        month_num = 7
    elif month == "08" or month == "Aug" or month == "August":
        folder = year + " 08 - August"
        month_num = 8
    elif month == "09" or month == "Sep" or month == "September":
        folder = year + " 09 - September"
        month_num = 9
    elif month == "10" or month == "Oct" or month == "October":
        folder = year + " 10 - October"
        month_num = 10
    elif month == "11" or month == "Nov" or month == "November":
        folder = year + " 11 - November"
        month_num = 11
    elif month == "12" or month == "Dec" or month == "December":
        folder = year + " 12 - December"
        month_num = 12
    else:
        raise NameError("Invalid Month")
    return (parent_folder + "\\" + folder, month_num)


def copy_folder_contents(source_loc, destination_loc):
    # function receives source and destination folder strings where both locations exist
    # iterates all files from source and copies them to destination

    # Get list of source files
    source_file_list = os.listdir(source_loc)

    # Loop through source files and copies to destination
    for file in source_file_list:
        current_file = source_loc + "\\" + file
        # print(current_file)
        shutil.copy(current_file, destination_loc)
        # print(os.listdir(destination_loc))


def get_meeting_date(year: int, month: int) -> str:
    cal = calendar.monthcalendar(year, month)
    first_week = cal[0]
    second_week = cal[1]
    third_week = cal[2]

    # If a Saturday presents in the first week, the second Saturday
    # is in the second week.  Otherwise, the second Saturday must
    # be in the third week.
    if first_week[calendar.TUESDAY]:
        return second_week[calendar.TUESDAY]
    else:
        return third_week[calendar.TUESDAY]


def rename_meeting_files(file_dir, meeting_date):
    file_list = os.listdir(file_dir)
    print(file_list)
    year = meeting_date.strftime("%Y")
    month = meeting_date.strftime("%m")
    day = meeting_date.strftime("%d")
    # print("Year: " + year + " Month: " + month + " Day: " + day)

    minutes_filename = file_dir + "\\" + "Minutes-Council 8291-"+ year + month + day + ".gdoc"
    slides_filename = file_dir + "\\" + "Meeting Slides - "+ year + "-" + month + ".pptx"

    print(minutes_filename)
    print(slides_filename)

    for file in file_list:
        spec_file = file_dir + "\\" + file
        if file.find("Minutes-Council") != -1:
            print("Minutes: " + spec_file)
            os.rename(spec_file, minutes_filename)
            print(os.listdir(file_dir))
        elif file.find("Slides") != -1:
            print("Slides: " + spec_file)
            os.rename(spec_file, slides_filename)
            print(os.listdir(file_dir))


def create_meeting(month: str, copy_month: str, parent_folder: str, year: str):
    # Receive month to create the meeting files for and the month folder to copy from
    # create new month folder
    # copy all files from source to new month folder

    # create the source and destination locations strings
    source_loc, source_month_num = month_folder_path(copy_month, parent_folder, year)
    destination_loc, month_num = month_folder_path(month, parent_folder, year)
    print("Source: " + source_loc)
    print("Destination: " + destination_loc)

    # Destination directory should not already exist, if yes: throw error, if no: create
    if os.path.isdir(destination_loc):
        raise NameError("Folder Already Exists!")
    else:
        # Create new folder
        os.mkdir(destination_loc)

    # Copy files to new folder
    copy_folder_contents(source_loc, destination_loc)

    # Get meeting date for file renaming
    meeting_date = date.datetime(int(year), month_num, int(get_meeting_date(int(year), month_num)))
    rename_meeting_files(destination_loc, meeting_date)