"""
    # Download PO copy from Oracle base on the PO list in excel
    # Ken: Aug 2019
"""

import pyautogui
import webbrowser
import time
import pandas as pd



url='https://wwwin-cg1prd.cisco.com:443/OA_HTML/RF.jsp?function_id=75&resp_id=51179&resp_appl_id=201&security_group_id=0&lang_code=US'

# 每一个步骤关键点坐标位置及颜色[(x,y),(R,G,B,A)]
step_setting = {'step_logon':[(813, 725), (4, 159, 217, 255),(1076, 383), (99, 178, 70, 255)],#(before mouse move to the x/y)logon button in user/pw window,'Call Me' button under Duo window
                'step_duo_recall':[(805, 594), (161, 54, 57, 255)], #if Duo not answer/confirmed, will recall again
                'step1':[(405,208), (45, 132, 197, 255)], #top bar (check readiness only).select 'single request' via 'enter'
			    'step2':[(360, 113), (45, 132, 197, 255)], #top bar (check readiness only). Input "Cisco printed PO report" (already focused) then 'enter'
				'step3':[(721, 238), (45, 132, 197, 255)], #top bar (check readiness only). Input 'PO_from' (already focused), then 'tab','tab','enter'
				'step4':[(360, 113), (45, 132, 197, 255)], #top bar (check readiness only). Use 'tab' to focus 'Delivery Opts' button, use 'enter' to next
                'step5':[(171, 165), (203, 216, 234, 255)],#email tab. Use 'ctrl+m' to select 'email' tab
                'step6':[(216, 214), (255, 238, 168, 255)],#'From' box. Select 'From' using mouseDown/Up;
				'step7':[(216, 214), (255, 238, 168, 255)],#'From' box after selection (check readiness). Input user then use 'tab' to switch to subject and TO/CC; use 'enter' to next
				'step8':[(211, 312), (255, 238, 168, 255)],#'To' box. Select using mouseDown/Up then input email, then use 'tab' to 'CC' and input another email.

				'step9':[(360, 113), (45, 132, 197, 255)], # x/y/color - as step2 (check readiness). press 'enter' to 'submit'
                'step10':[(722, 390), (45, 132, 197, 255)], # x/y/color: top bar (check readiness only). Decide if want to submit another PO
					 }

def logon_and_duo(step='step_logon'):
    '''
    Use this to check if logon/security check window pop up, and action accordingly.
    :param step:
    :return:
    '''

    pos1_color,pos2_color,pos3_color = (1, 1, 1, 1),(1, 1, 1, 1),(1, 1, 1, 1)  # set initial base color
    #below check all the possible starting window: Logon/Duo/Oracle start window (so it won't stuck if not found)
    while pos1_color != step_setting.get(step)[1] and pos2_color != step_setting.get(step)[3] and pos3_color != step_setting.get('step1')[1]:  # 如果多条件是用 AND
        time.sleep(2)

        screen = pyautogui.screenshot().resize(pyautogui.size())
        pos1_color = screen.getpixel(step_setting.get(step)[0])
        pos2_color = screen.getpixel(step_setting.get(step)[2])
        pos3_color = screen.getpixel(step_setting.get('step1')[0])
        print('looking for Logon/Duo/start window...')

    if pos1_color == step_setting.get(step)[1]: #if it's logon window
        pyautogui.moveTo(step_setting.get(step)[0],duration=0.2)
        pyautogui.click()
    elif pos2_color == step_setting.get(step)[3]: # or if it's Dup window
        pyautogui.moveTo(step_setting.get(step)[2], duration=0.2)
        pyautogui.click()
    else: # or if it's already Oracle start window
        pass


def duo_recall(step='step_duo_recall',wait_time_for_next_call=10):
    '''
    If Duo not answered/confirmed, will recall this again for 3 times .
    :param step:
    :return:
    '''
    need_recall=False

    pos1_color,pos2_color =(1, 1, 1, 1),(1, 1, 1, 1)  # set initial base color
    #below check all the possible starting window: Logon/Duo/Oracle start window (so it won't stuck if not found)
    while pos1_color != step_setting.get(step)[1] and pos2_color != step_setting.get('step1')[1]:  # 如果多条件是用 AND
        time.sleep(2)

        screen = pyautogui.screenshot().resize(pyautogui.size())
        pos1_color = screen.getpixel(step_setting.get(step)[0])
        pos2_color = screen.getpixel(step_setting.get('step1')[0])
        print('looking for Duo or start window...')

    if pos1_color == step_setting.get(step)[1]: # Duo window
        time.sleep(wait_time_for_next_call)
        pyautogui.moveTo(step_setting.get('step_logon')[2], duration=0.2)
        pyautogui.click()
    else: # or if it's already Oracle start window
        pass


def readiness_check(step=None):
    '''
    common function to check if a window or a position is ready&available by checking it's color
    :param step:link to steps in step_setting
    :return: True when it's ready
    '''
    pyautogui.moveTo(step_setting.get(step)[0],duration=0.3)
    pos_color = (1, 1, 1, 1)  # set initial base color
    while pos_color != step_setting.get(step)[1]:  # 如果多条件是用 OR
        time.sleep(0.5)

        screen = pyautogui.screenshot().resize(pyautogui.size())
        pos_color = screen.getpixel(step_setting.get(step)[0])

    return pos_color


def step1_single_request_ok():
    '''
    First oracle window: select 'single request'
    :return: None
    '''
    # just hit 'enter' since OK button is focused already
    pyautogui.press('enter')

def step2_start_input(msg='Cisco printed PO report'):
    '''
    Input 'Cisco printed PO report' (box already focused), then Enter
    :param msg:  'Cisco printed PO report'
    :return: None
    '''

    # just type the words since box already focused
    pyautogui.typewrite(message=msg,interval=0.03)
    pyautogui.press('enter')


def step3_enter_po(po=None):
    '''
    Input PO number and selection (ALL)
    :param po: PO_number
    :return: None
    '''
    # just typewrite since box already focused
    pyautogui.typewrite(message=po,interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.5)

    pyautogui.press('tab')
    pyautogui.typewrite(message='ALL',interval=0.03)

    pyautogui.press('enter')

def step4_click_delivery_opt_button(step='step4'):
    '''
    Select the 'delivery opts' option
    :param step: 'step4'
    :return: None
    '''
    #pyautogui.moveTo(step_setting.get(step)[2])
    #pyautogui.click()
    #pyautogui.moveTo((0,0),duration=1)

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')


def step5_select_email_tab(step='step5'):
    '''
    Select the 'email' tab
    :param step:
    :return:
    '''
    pyautogui.moveTo(step_setting.get(step)[0], duration=0.2)
    pyautogui.hotkey('ctrl','m')

def step6_click_from_box(step='step6'):
    '''
    Click the From box after selecting email tab
    :param step:
    :return:
    '''

    pyautogui.moveTo(step_setting.get('step6')[0],duration=0.2)
    #pyautogui.click()
    pyautogui.mouseDown()
    pyautogui.mouseUp(duration=1)


def step7_input_delivery_opt(input_record=None):
    '''
    Input the delivery options including username, subject and emails addresses, then 'enter'
    :param step: 'step6'
    :param input_record: values from the input_df
    :return: None
    '''
    po=str(input_record[0])
    user=input_record[1]

    pyautogui.typewrite(message=user,interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.5)

    pyautogui.typewrite(message=po,interval=0.03)
    time.sleep(0.5)


def step8_input_email(step='step8',input_record=None):
    '''
    Select the To/CC and put in email addressess
    :param step:
    :return:
    '''

    email_1=input_record[2]
    email_2 = input_record[3]

    pyautogui.moveTo(step_setting.get(step)[0], duration=0.2)
    pyautogui.mouseDown()
    pyautogui.mouseUp(duration=1)

    pyautogui.typewrite(message=email_1, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.typewrite(message=email_2,interval=0.03)
    time.sleep(0.5)
    pyautogui.press('enter')


def step9_submit(step='step9'):
    '''
    Press enter to submit this PO
    :param step: 'step7'
    :return: None
    '''

    pyautogui.press('enter')


def step10_decide_if_continue(next_po=True):
    '''
    Decide if need to submit new PO
    :param next_po: boolean to continue or quit
    :return:
    '''

    if next_po==True:
        pyautogui.press('enter')
    else:
        pyautogui.hotkey('command','q') #退出oracle
        time.sleep(3)
        pyautogui.hotkey('command', 'q') #退出firefox


def read_data(fname):
    email=[]
    df = pd.read_excel(fname)
    po = df.PO_NUMBER.values
    user = df.USER_NAME.values
    email.append(df.EMAIL_1.values)
    email.append(df.EMAIL_2.values)

    return df

def po_download(input_df):
    '''
    Based on df that contains PO/user/email to download the PO copy from Oracle
    :param df: dataframe with header row: PO_NUMBER, USER_NAME,EMAIL
    :return: ?
    '''

    #step1 只需在开始时做一次
    readiness_check(step='step1')
    step1_single_request_ok()

    print('Start downloading PO...')

    loaded_po=[]
    for i, record in enumerate(input_df.values):
        print(record)

        readiness_check(step='step2')
        step2_start_input(msg='Cisco printed PO report')

        readiness_check(step='step3')
        po_number=str(record[0])
        step3_enter_po(po=po_number)

        readiness_check(step='step4')
        step4_click_delivery_opt_button(step='step4')

        readiness_check(step='step5')
        step5_select_email_tab(step='step5')
        step6_click_from_box(step='step6')

        readiness_check(step='step7')
        step7_input_delivery_opt(input_record=record)
        step8_input_email(step='step8', input_record=record)

        readiness_check(step='step9')
        step9_submit(step='step9')

        readiness_check(step='step10')
        if i==input_df.shape[0]-1:
            continue_next_po=False
        else:
            continue_next_po = True
        step10_decide_if_continue(next_po=continue_next_po)

        loaded_po.append(record)

    return loaded_po


if __name__=='__main__':
    pyautogui.FAILSAFE = True

    #从文件读取数据
    input_df=read_data(fname='/users/wangken/py/auto/po_download/PO.xlsx')

    # 打开Oracle
    webbrowser.open(url)

    time.sleep(1)
    #first on logon window
    logon_and_duo(step='step_logon')
    # second on Duo window
    logon_and_duo(step='step_logon')

    #recall again is Duo is not answered/confirmed
    duo_recall(step='step_duo_recall',wait_time_for_next_call=10)

    # 等待15s后开始循环判断界面是否准备好了
    #time.sleep(15)

    #开始下载PO
    loaded_po=po_download(input_df)
    po_list=[x[0] for x in loaded_po]

    #print('Below PO have executed PO download{}:\n{}'.format(len(po_list),po_list))
