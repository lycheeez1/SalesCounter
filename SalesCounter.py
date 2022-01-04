import os, time
from tkinter import *
import tkinter.ttk as ttk
import tkinter.messagebox as mb
from tkinter import filedialog
import concat_dataframe as ccdf
#import threading


def open_file():
    ent.delete(0, "end")
    root.filename = filedialog.askopenfilename(initialdir=r'', title="파일 불러오기")
    path = root.filename
    ent.insert(0, path)
#    filename = os.path.basename(path)
#    ent.insert(0, filename)
    print(path)


def open_manual():
    mb.showinfo("Manual",
                '''
1. [파일선택] 버튼 클릭 → 주문내역 엑셀 파일 선택
2. [실행] 버튼 클릭 → 판매량 엑셀 파일이 생성됨
※ 결과 파일 생성 위치 : 공용PC/메인컴퓨터/물류팀/입고파트/발주 재고리스트/판매량

''')


# [실행] 함수
def execute():
    path = ent.get()
    try:
        if path == '':
            mb.showwarning("경고", "엑셀 파일을 불러오세요")
        else:
            progressbar.place(relx=0.155, rely=0.8)
            ccdf.concat_all(path)
#            cp.cal_coupang(path)        # 수량 계산!

            for i in range(1, 101):
                time.sleep(0.001)
                pg_var.set(i)
                progressbar.update()

            time.sleep(0.5)
            response = mb.showinfo("성공", "다운 완료")
            if response != 1:
                progressbar.place_forget()

    except Exception as e:
        mb.showwarning("경고", e)
        print(str(e))
        ent.delete(0, "end")    # 엔트리 내용 비우기
        # 에러 로그 남기기



if __name__ == "__main__":
    # 루트 윈도우
    root = Tk()
    root.title("sales counter")
    root.resizable(False, False)    # 창 크기 고정 (가로, 세로)

    win_w, win_h = 370, 240
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = (screen_w - win_w) / 3
    y = (screen_h - win_h) / 4
    root.geometry("%dx%d+%d+%d" %(win_w, win_h, x, y))      # 창 크기
    root.configure(background="ghostwhite")

    # 메뉴
    menu = Menu(root)
    root.config(menu=menu)
    menu_help = Menu(menu, tearoff=0)
    menu_help.add_command(label="사용 매뉴얼", command=open_manual)  # 서브메뉴
    menu.add_cascade(label="Help", menu=menu_help)

    # 라벨 : 안내 메시지
    lb = Label(root, text="PC 사양에 따라 결과를 산출하는데 1~3분 소요됩니다.\n CPU 과부하로 인해 (응답없음) 상태가 될 수 있으나 \n   정상 실행 중인 것이니 프로그램을 종료하지 말고 기다려주세요.   ")
    lb.place(x=0, y=0)
    
    # 엔트리(한줄 입력창)
    ent = Entry(root, width=26, relief="ridge")
    ent.place(x=45, y=70)

    # [파일 선택] 버튼
    btn_open = Button(root, text="파일 선택", command=open_file,
                      bg="whitesmoke", bd=1, padx=10, activebackground="lavender", cursor="hand2")
    btn_open.place(x=238, y=66)

    # [실행] 버튼
    btn = Button(root, text="실행", command=execute,
             font=9, padx=90, pady=10, bd=1,
             cursor="hand2", overrelief="sunken", activebackground="lavender", activeforeground="black")
    btn.place(x=70, y=110)

    # 프로그레스바
    pg_var = DoubleVar()
    progressbar = ttk.Progressbar(root, maximum=100, length=250, variable=pg_var)



    root.mainloop()     # 윈도우가 종료 버튼을 누르기 전까지 종료되지 않도록 함
