import Bot_Functions as bot
import datetime as date

output_file = "C:\Python37\WorkingFiles\contact_list.txt"
contact_import = 'C:\Python37\WorkingFiles\ContactExtract.csv'
meeting_folder = "D:\Google MyDrive\\0 - KofC Materials\Meetings"
test_folder = "D:\Google MyDrive\\0 - KofC Materials\Meetings\\2020 03 - March"
bot.update_contacts(contact_import, output_file)

#bot.create_meeting("March", "February", meeting_folder, "2020")

# print(bot.get_meeting_date(2020,3))
#x = date.datetime(2020,3,10)
#bot.rename_meeting_files(test_folder, x)

#bot.update_contacts(contact_import, working_directory)



