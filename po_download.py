"""
    # Download PO copy from Oracle base on the PO list in smartsheet
    # Ken: Aug 2019; updated Jun, 2020
"""

import pyautogui
import webbrowser
import time
import pandas as pd
from smartsheet_handler import SmartSheetClient
import smartsheet



url='https://wwwin-cg1prd.cisco.com:443/OA_HTML/RF.jsp?function_id=75&resp_id=51179&resp_appl_id=201&security_group_id=0&lang_code=US'

# 每一个步骤关键点坐标位置及颜色[(x,y),(R,G,B,A)]
step_setting = {'step_logon':[(813, 725), (4, 159, 217, 255),(1076, 383), (99, 178, 70, 255)],#(before mouse move to the x/y)logon button in user/pw window,'Call Me' button under Duo window
                'step_duo_recall':[(805, 594), (161, 54, 57, 255)], #if Duo not answer/confirmed, will recall again
                'step1':[(405,208), (45, 132, 197, 255)], #start option window top bar.select 'single request' via 'enter'
			    'step2':[(360, 113), (45, 132, 197, 255)], #main window top bar. Name is focused, input "Cisco printed PO report" and so on
				'step3':[(578, 208), (45, 132, 197, 255)], # PO input window top bar. PO is focused, input po and other info
				'step4':[(468, 567), (211, 228, 244, 255)], # submit button (use step1 loc to check readiness)
                #'step4':[(531, 602), (106, 86, 59, 255)], # input employee id; then tab, tab, input email address; then enter to exit the window
				#'step9':[(360, 113), (45, 132, 197, 255)], # x/y/color - as step2 (check readiness). press 'enter' to 'submit'
                'step5':[(722, 390), (45, 132, 197, 255)], # Decide if want to submit another PO
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


def readiness_check(step='stepx'):
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

def get_pos_color(step='stepx'):
    '''
    Get position color
    :param step:link to steps in step_setting
    :return: True when it's ready
    '''
    pyautogui.moveTo(step_setting.get(step)[0],duration=0.3)
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

def step2_basic_option_input(po_number,msg='Cisco Email printed PO report'):
    '''
    Input 'Cisco printed PO report' (box already focused), then Enter
    :param msg:  'Cisco printed PO report'
    :return: None
    '''
    if po_number[:3]=='555':
        op_unit='NETHERLANDS Operating'
    elif po_number[:3]=='202':
        op_unit = 'CISCO US OPERATING UNIT'
    else:
        op_unit = 'NETHERLANDS Operating'

    # just type the words since box already focused
    pyautogui.typewrite(message=msg,interval=0.03)
    pyautogui.press('tab') # this tab move to po window by deafault, with op_unit as NETHERLANDS Operating by default

    time.sleep(2)
    step='step3'
    pos_color=get_pos_color(step=step)
    if pos_color!=step_setting[step][1]: # 如果没有跳转到po_window
        pyautogui.typewrite(message=op_unit, interval=0.03)
        pyautogui.press('tab') # move to po window

def step3_enter_po_id_email(po_number,employee_id,email):
    '''
    Input PO number and selection (ALL)
    :param po: PO_number
    :return: None
    '''
    # just typewrite since box already focused
    pyautogui.typewrite(message=po_number,interval=0.03)
    pyautogui.press('tab')
    pyautogui.typewrite(message=po_number, interval=0.03)
    time.sleep(0.5)
    pyautogui.press('tab')
    pyautogui.typewrite(message='All',interval=0.03)

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.typewrite(message=employee_id, interval=0.03)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.typewrite(message=email, interval=0.03)

    pyautogui.press('tab')
    pyautogui.press('enter')


def step4_submit():
    '''
    Press enter to submit this PO
    :param step: 'step7'
    :return: None
    '''
    pyautogui.moveTo(step_setting.get('step4')[0], duration=0.2)
    pyautogui.click()


def step5_decide_if_continue(next_po=True):
    '''
    Decide if need to submit new PO
    :param next_po: boolean to continue or quit
    :return:
    '''

    if next_po==True:
        pyautogui.press('enter')
    else:
        pyautogui.hotkey('command','q') #退出oracle
        #time.sleep(3)
        #pyautogui.hotkey('command', 'q') #退出firefox



def po_download(smartsheet_client, sheet_id, df_po):
    '''
    Based on df that contains PO/user/email to download the PO copy from Oracle
    :param df: data frame with header row: PO_NUMBER, USER_NAME,EMAIL
    :return:
    '''

    #step1 option window 只需在开始时做一次
    readiness_check(step='step1')
    step1_single_request_ok()

    print('Start downloading PO...')

    po_downloaded=[]
    i=1
    for row in df_po.itertuples(index=False):
        po_number=str(row.PO_NUMBER).strip()
        employee_id=str(row.EMPLOYEE_ID).strip()
        email=str(row.CM_EMAIL).strip() +', ' + row.USER_NAME.strip() + '@cisco.com'
        #email = str(row.CM_EMAIL).strip()

        readiness_check(step='step2')
        step2_basic_option_input(po_number,msg='Cisco Email Printed PO Report')

        readiness_check(step='step3') # PO window
        step3_enter_po_id_email(po_number,employee_id,email)

        readiness_check(step='step2') #back to step2 window
        step4_submit() # submit button

        if i==df_po.shape[0]:
            continue_next_po=False
        else:
            continue_next_po = True
            i+=1

        po_downloaded.append(po_number)
        update_po_status_in_smartsheet(smartsheet_client, sheet_id, df_po, po_number)

        readiness_check(step='step5')
        step5_decide_if_continue(next_po=continue_next_po)

    return po_downloaded

def read_po_from_smartsheet(smartsheet_client,sheet_id):
    '''
    Read PO data from smartsheet
    :return:
    '''
    # 从smartsheet读取po data
    df = smartsheet_client.get_sheet_as_df(sheet_id, add_row_id=True, add_att_id=False)

    df=df[df.STATUS!='DOWNLOADED']

    return df

def update_po_status_in_smartsheet(smartsheet_client,sheet_id,df_po,po_number):
    """
    # 更新smartsheet
    :param smartsheet_client:
    :param sheet_id:
    :param df:
    :param po_number:
    :return:
    """
    row_id=int(df_po[df_po.PO_NUMBER==po_number].iloc[0,-1])
    smartsheet_client.update_row_with_dict(ss=smartsheet.Smartsheet(token), process_type='update',
                                           sheet_id=sheet_id,
                                           row_id=row_id,
                                           update_dict=[{'STATUS': 'DOWNLOADED','DATE':pd.Timestamp.today().strftime('%Y-%m-%d')}])

if __name__=='__main__':
    pyautogui.FAILSAFE = True

    token = '5wjf4z6dw5s51urk0cao7vkxok'
    sheet_id = '6111099402643332'
    proxies = None  # for proxy server
    smartsheet_client = SmartSheetClient(token, proxies)

    #从文件读取数据
    #input_df=read_data(fname='/users/wangken/py/auto/po_download/PO.xlsx')
    df_po=read_po_from_smartsheet(smartsheet_client,sheet_id)
    df_po=df_po[df_po.STATUS!='DOWNLOADED']
    if df_po.shape[0]==0:
        print("\n**All PO in smartsheet marked as 'DOWNLOADED' already. Program stop.")
    else:
        print(df_po)
        # 打开Oracle
        webbrowser.open(url)
        # 等待30s后开始循环判断界面是否准备好了
        time.sleep(30)

        """ discard the autologon steps and make it manually controlled
        time.sleep(1)
        #first on logon window
        logon_and_duo(step='step_logon')
        # second on Duo window
        logon_and_duo(step='step_logon')
    
        #recall again is Duo is not answered/confirmed
        duo_recall(step='step_duo_recall',wait_time_for_next_call=10)
        """


        #开始下载PO
        po_downloaded=po_download(smartsheet_client, sheet_id, df_po)

        print('{} PO have been download:'.format(len(po_downloaded)))
        for po in po_downloaded:
            print(po)
