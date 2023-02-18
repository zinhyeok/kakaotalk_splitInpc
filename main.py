#항상 실행 시 X -> 명령어를 확인해야함
#자동 실행도 시간 확인 하려면 항상 실행해야함
# 실행하면 당일에 아침 5:30~9:00, 9:00~10:00까지 채팅리스트를 확인 -> string으로 반환
# -> 반환 형식: n월 n일 지각자: ~ n월 n일 기절자: ~ 없으면 출력 안함
# 시간은 변경 가능하게끔 설정


#pip install https://wonpaper.tistory.com/375
import time, win32con, win32api, win32gui, ctypes
import requests
from pywinauto import clipboard # 채팅창내용 가져오기 위해
import pandas as pd # 가져온 채팅내용 DF로 쓸거라
# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = '투자'

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con


## 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RichEdit50W", None)
    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

# # 채팅내용 가져오기
def copy_chatroom(chatroom_name):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndListControl = win32gui.FindWindowEx(hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    ##본문을 클립보드에 복사 ( ctl + c , v )
    PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(1)
    PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    ctext = clipboard.GetData()
    # print(ctext)
    return ctext


# 조합키 쓰기 위해
def PostKeyEx(hwnd, key, shift, specialkey):
    if IsWindow(hwnd):

        ThreadId = GetWindowThreadProcessId(hwnd, None)

        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)


## 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


# # 채팅방 열기
def open_chatroom(chatroom_name):
    # # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)    # ㄴ시작핸들을 첫번째 자식 핸들(친구목록) 을 줌(hwndkakao_edit2_1)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)   # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)


# # 채팅내용 하루치 저장 df로 반환
def chat_last_save():
    open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    ttext = copy_chatroom(kakao_opentalk_name)  # 채팅내용 가져오기
    a = ttext.split('\r\n')   # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)    # DF 으로 바꾸기
    cur_year = str(time.localtime().tm_year) #오늘 날짜 불러오기
    cur_month = str(time.localtime().tm_mon)
    cur_day = str(time.localtime().tm_mday)

    # 오늘 날짜와 일치하는 부분의 index 반환하기
    index_num = df.index[df[0].str.contains(cur_year+"년" + " " + cur_month+"월"+" "+cur_day+"일")]
    index_num = index_num.tolist()[0]
    return df[index_num:]
    # df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')  # 정규식으로 채팅내용만 남기기
    # return df.index[-2], df.iloc[-2, 0]

#하루치 채팅 중 미션으로 시작하는 채팅부터만 출력
def chat_misson_startStr(word):
    word = str(word)
    df = chat_last_save()
    index_num = df.index[df[0].str.contains(word)]
    index_num = int(index_num.tolist()[0])
    return df.loc[index_num:]

#하루치 채팅 내용에서 사람 출력(정해놓은 시간대까지만), return 3개, 기상미션 성공, 지각, 기절
# 이부분 함수만 건드는 것을 추천!

def chat_misson_Success(lst, start, late, out):
    df = chat_misson_startStr(str(start))
    late_index = df.index[df[0].str.contains(str(late))]
    late_index = int(late_index.tolist()[0])

    out_index = df.index[df[0].str.contains(str(out))]
    out_index = int(out_index.tolist()[0])

    df_suc = df.loc[:late_index]
    df_late = df.loc[late_index:out_index]

    lst_all = lst
    lst_suc = []
    lst_late= []

    for i in df_suc.index:
        name = list(df_suc.loc[i][0].split("]"))
        lst_suc.append(name[0].lstrip("["))

    lst_suc = list(set(lst_suc))


    for i in df_late.index:
        name = list(df_late.loc[i][0].split("]"))
        lst_late.append(name[0].lstrip("["))

    lst_late = list(set(lst_late))
    lst_late = [x for x in lst_late if x not in lst_suc]

    lst_out = list(set([x for x in lst_all if x not in lst_suc+lst_late]))


    return lst_suc, lst_late, lst_out


def main():
    #lst_all에 단톡방에 있는 전 인원의 이름을 list, string type으로 넣기
    lst_all = []
    suc, late, out = chat_misson_Success(lst_all,"미션","지각컷","기절컷")
    print(suc, late, out)
    #오늘 날짜 불러오기
    cur_month = str(time.localtime().tm_mon)
    cur_day = str(time.localtime().tm_mday)
    text = "{}월 {}일 기상미션 결과 \n 성공: {} \n 지각: {} \n 기절: {}".format(cur_month, cur_day, suc, late, out)
    kakao_sendtext(kakao_opentalk_name, text)

if __name__ == '__main__':
    main()