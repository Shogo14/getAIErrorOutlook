from O365 import Account
import file_make
import time
import constants

start_time = time.time()

credentials = (constants.CLIENT_ID, constants.CLIENT_SECRET)
account = Account(credentials)
limit_count = constants.LIMIT_COUNT

mailbox = account.mailbox()
inbox = mailbox.inbox_folder()
#terget foloder
ncol_box = inbox.get_folder(folder_name='0.NCOL')
ai_error_box = ncol_box.get_folder(folder_name='AIエラー')
un_read_messages = [message for message in ai_error_box.get_messages(limit=limit_count ) if not message._Message__is_read]
if 1 > len(un_read_messages):
    print("un read message is not exist!")
    print('Execute Time: ' + str(round(time.time() - start_time, 2)) + ' seconds')
    exit(0)
print(len(un_read_messages))
data_lists = []
for message in un_read_messages:
    read_flag = message._Message__is_read
    if not read_flag:
        body = message._Message__body_preview
        errorType = body[body.find('errorType:')+11:body.find('errorType:')+12]
        errorTime = body[body.find('errorTime:')+11:body.find('errorTime:')+25]
        session_id = body[body.find('session_id:')+12:body.find('session_id:')+24]
        received_date = message._Message__received
        if "session_id:" in errorTime:
            errorTime = ""
        data_dict = {"errorType": errorType,"errorTime" : errorTime,"session_id": session_id,"received_date": received_date}
        data_lists.append(data_dict.copy())
        message.mark_as_read()

if len(data_lists) > 0:
    file_make.write_excel(data_lists, constants.FILE_NAME)

print('End. unread message : '+ str(len(data_lists)))
print('Execute Time: ' + str(round(time.time() - start_time, 2)) + ' seconds')

