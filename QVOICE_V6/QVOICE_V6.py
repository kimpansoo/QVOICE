from pynput.keyboard import Controller
import os
import threading
import speech_recognition as sr
import serial
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import psutil
import sys
import glob

speech = []
speech_1_on_port = []
data_fin = []
backup_list_port = []
backup_open_list = []
backup_open_list_port = []


def serial_ports(): #사용 가능한 통신포트 찾기
    global result
    if sys.platform.startswith('win'):
        ports = ['COM%s' % (i + 1) for i in range(256)]
    elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
        # this excludes your current terminal "/dev/tty"
        ports = glob.glob('/dev/tty[A-Za-z]*')
    elif sys.platform.startswith('darwin'):
        ports = glob.glob('/dev/tty.*')
    else:
        raise EnvironmentError('Unsupported platform')

    result = []
    for port in ports:
        try:
            s = serial.Serial(port)
            s.close()
            result.append(port)
        except (OSError, serial.SerialException):
            pass
    return result


if __name__ == '__main__':
    result = serial_ports()
   #print(result)


class voiceloop(threading.Thread):

    global backup_list
    mykeyboard = Controller()

    def run(self) -> None:

        while True:
            voice = self.CollectVoice()

            if voice != False and myThread.rflag == True:
                print(voice)
                self.Pasting(voice)

            if myThread.rflag == False:
                break


    def Pasting(self, myvoice):
        for character in myvoice:
            self.mykeyboard.type(character)
        self.mykeyboard.type(" ")

    def CollectVoice(self):

        # get microphone device on notebook or desk top
        listener = sr.Recognizer()
        voice_data = ""

        with sr.Microphone() as raw_voice:

            try:
                img_frm.config(image=mic3_img)
                print("Adjusting")
                listener.adjust_for_ambient_noise(raw_voice)

                # adjust setting values
                listener.dynamic_energy_adjustment_damping = 0.2
                listener.pause_threshold = 0.6
                listener.energy_threshold = 600
                img_frm.config(image=mic1_img)
                print("Say something!")
                audio = listener.listen(raw_voice)
                img_frm.config(image=mic2_img)
                voice_data = listener.recognize_google(audio, language='ko')
                voice_label.config(text=voice_data)
                data_back()
                for i in range(len(backup_list)):
                    #print(backup_list)
                    #print(len(backup_list))
                    try:
                        if (voice_data == backup_list[i][0]):  #프로그램 시작 함수
                            filepath = backup_list[i][1]
                            os.startfile(filepath)
                        elif (voice_data == backup_list[i][3]):  #프로그램 종료 함수
                            s = (backup_list[i][2]).lstrip('[').rstrip(']')
                            s1 = s.lstrip("'").rstrip("'")
                            backup_list[i][2] = s1
                            for proc in psutil.process_iter():
                                if proc.name() == backup_list[i][2]:
                                    proc.kill()
                            s=''
                            s1=''
                    except:
                        pass

                b=Combx1.get()
                #print(b)
                ser = serial.Serial("%s"%b, 9600, timeout=1)

                for i in range(len(backup_open_list_port)):
                    if (voice_data == backup_open_list_port[i][0]):  # 켜기
                        ser.write(str.encode("on%s" %i))
                    elif (voice_data == backup_open_list_port[i][1]):  # 꺼기
                        ser.write(str.encode("off%s" %i))
                if (voice_data == backup_open_list_port[0][0]):label2.config(foreground="red")
                elif (voice_data == backup_open_list_port[0][1]):label2.config(foreground="gray15")
                elif (voice_data == backup_open_list_port[1][0]):label3.config(foreground="red")
                elif (voice_data == backup_open_list_port[1][1]):label3.config(foreground="gray15")
                elif (voice_data == backup_open_list_port[2][0]):label4.config(foreground="red")
                elif (voice_data == backup_open_list_port[2][1]):label4.config(foreground="gray15")
                elif (voice_data == backup_open_list_port[3][0]):label5.config(foreground="red")
                elif (voice_data == backup_open_list_port[3][1]):label5.config(foreground="gray15")
                
                elif (voice_data == "선풍기 1단"):
                    ser.write('A'.encode())
                    label6.config(foreground="red")
                elif (voice_data == "선풍기 2단"):
                    ser.write('B'.encode())
                    label6.config(foreground="red")
                elif (voice_data == "선풍기 3단"):
                    ser.write('C'.encode())
                    label6.config(foreground="red")
                elif (voice_data == "선풍기 4단"):
                    ser.write('D'.encode())
                    label6.config(foreground="red")
                elif (voice_data == "선풍기 회전"):
                    ser.write('E'.encode())
                    label6.config(foreground="red")
                elif (voice_data == "선풍기 켜"):
                    ser.write('F'.encode())
                    label6.config(foreground="red")
                elif (voice_data == "선풍기 꺼"):
                    ser.write('G'.encode())
                    label6.config(foreground="gray15")
                elif (voice_data == "PC 종료"):
                    os.system("shutdown -s -t 0")
                elif (voice_data == "PC 재시작"):
                    os.system("shutdown -r -t 0")


                ser.close()

            except UnboundLocalError:
                pass

            except sr.UnknownValueError:
                print("could not understand audio")
                return False

            return str(voice_data)

# tk 함수 모음 ----------------
def data_fopen(): #파일찾기

    file_entry.delete("1.0", END)
    global filename
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=(("all files", "*.*"), ("PPTX files", "*.pptx")))
    file_entry.insert(INSERT, filename)
    if (btn_voice['state'] == tk.DISABLED):
        btn_voice['state'] = tk.NORMAL
        #btn_back['state'] = tk.DISABLED
    speech_entry.focus_set()
    if (btn_del['state'] == tk.DISABLED or btn_list_del['state'] == tk.DISABLED ):
        btn_del['state'] = tk.NORMAL
        btn_list_del['state'] = tk.NORMAL

def data_voice_save(): #음성입력 및 저장

    speech = speech_entry.get()
    if len(speech) > 1:
        pass
    else:
        messagebox.showinfo(title="경고", message="두음절 이상 입력 해주세요")
    files_0 = file_entry.get('1.0', END)
    files = files_0.replace("\n", "")
    speech_0 = speech_entry.get()
    speech = speech_0.rstrip()
    speech_off_0 = speech_off_entry.get()
    speech_off = speech_off_0.rstrip()

    data = files.split("/")
    data_fin = [word for word in data if '.' in word]  # '.'가 포함된 모든 문자열 추출

    if len(files) < 3 or len(speech_0) < 2:
        messagebox.showinfo(title="저장실패", message="위 두칸 모두 입력해주세요")
    else:
        with open("back_up0.txt", "a") as file:
            file.write(f"{speech}-{files}-{data_fin}-{speech_off}\n")
        file_entry.delete("1.0", END)
        speech_entry.delete(0, END)
        speech_off_entry.delete(0, END)
        root.focus_set()

        btn_voice['state'] = tk.DISABLED

        data_back()
        data_list_open()


def data_back():  #적용 및 백업하기
    global backup_list
    backup_list=[]
    backup_list0=[]
    backup_list1 = []
    y =[]

    with open("back_up0.txt", "r") as file:
        lines = file.readlines()
        lines = [line.rstrip('\n') for line in lines]
        for i in range(len(lines)):
            backup_list0.insert(i, lines[i])
            backup_list1 = backup_list0[i].split('-')
            y.insert(0,backup_list1[0])
            backup_list.insert(i, backup_list1)

    x=len(backup_list)
    lb_content.config(text="{}".format(y))
    lb_count.config(text="[총{}개]".format(x))


def data_list_open():  #목록보기
    global lst
    lst = []
    lb.delete(0, END)
    f = open("back_up0.txt", "r")
    for x in f:
        lb.insert(END, x)
        lst.insert(0, x)
    f.close()

    if (btn_del['state'] == tk.DISABLED or btn_list_del['state'] == tk.DISABLED):
        btn_del['state'] = tk.NORMAL
        btn_list_del['state'] = tk.NORMAL


def data_list_del(): #리스트 선택요소 삭제하기
    global backup_list
    backup_list=[]
    s= lb.get(ANCHOR)
    sel = lb.curselection()
    lb.delete(sel[0])
    find_str = s #리스트 박스 선택한 요소 str
    while (find_str in lst):
        lst.remove(find_str)
    with open("back_up0.txt", "w") as file:
        file.write("")
    with open("back_up0.txt", "w") as file:
        file.write(''.join(lst))
    #backup_list=[]
    data_back()
    root.focus_set()

    data_list_open()

def data_del(): #삭제하기
    global backup_list
    MsgBox = tk.messagebox.askquestion("삭제확인창", "일괄삭제 할까요?")
    if MsgBox == 'yes':
        with open("back_up0.txt", "w") as file:
            file.write("")
        backup_list = []
        lb.delete(0,END)
    else:
        tk.messagebox.showinfo('삭제확인창', '삭제가 취소되었습니다.')


def help():
    lb.delete(0, END)
    with open("help.txt", "r", encoding="UTF-8") as file:
        for x in file:
           lb.insert(END, x)
    if (btn_del['state'] == tk.NORMAL or btn_list_del['state'] == tk.NORMAL ):
        btn_del['state'] = tk.DISABLED
        btn_list_del['state'] = tk.DISABLED


def data_voice_save_port():  # 외부 제어 음성입력 및 저장
    with open("back_up_port.txt", "w") as file:
        file.write("")

    speech_1_on_port = con1_on.get()
    speech_1_off_port = con1_off.get()
    speech_2_on_port = con2_on.get()
    speech_2_off_port = con2_off.get()
    speech_3_on_port = con3_on.get()
    speech_3_off_port = con3_off.get()
    speech_4_on_port = con4_on.get()
    speech_4_off_port = con4_off.get()

    speech_1_on_port = speech_1_on_port.rstrip()
    speech_1_off_port = speech_1_off_port.rstrip()
    speech_2_on_port = speech_2_on_port.rstrip()
    speech_2_off_port = speech_2_off_port.rstrip()
    speech_3_on_port = speech_3_on_port.rstrip()
    speech_3_off_port = speech_3_off_port.rstrip()
    speech_4_on_port = speech_4_on_port.rstrip()
    speech_4_off_port = speech_4_off_port.rstrip()

    if len(speech_1_on_port) < 2 or len(speech_1_off_port) <2:
        messagebox.showinfo(title="저장실패", message="두음절 이상 입력해주세요")
    else:
        with open("back_up_port.txt", "w") as file:
            file.write(f"{speech_1_on_port}-{speech_1_off_port}\n")
            file.write(f"{speech_2_on_port}-{speech_2_off_port}\n")
            file.write(f"{speech_3_on_port}-{speech_3_off_port}\n")
            file.write(f"{speech_4_on_port}-{speech_4_off_port}\n")

        con1_on.delete(0, END)
        con2_on.delete(0, END)
        con3_on.delete(0, END)
        con4_on.delete(0, END)
        con1_off.delete(0, END)
        con2_off.delete(0, END)
        con3_off.delete(0, END)
        con4_off.delete(0, END)
        root.focus_set()


        #data_back_port()
        data_list_open_port()


def data_list_open_port():  #외부 제어 목록보기
    global backup_open_list_port
    global lst2_port

    lst_port = []
    lst1_port = []
    lst2_port = []
    backup_open_list_port = []
    with open("back_up_port.txt", "r") as file:
        lines_port = file.readlines()
        lines_port = [line.rstrip('\n') for line in lines_port]
        for i in range(len(lines_port)):
            lst_port.insert(i, lines_port[i])
            lst1_port = lst_port[i].split('-')
            #print("lst1_port : ", lst1_port)
            backup_open_list_port.insert(i,lst1_port)
            #print("backup_open_list_port : ", backup_open_list_port)

        con1_on.insert(INSERT, backup_open_list_port[0][0])
        con2_on.insert(INSERT, backup_open_list_port[1][0])
        con3_on.insert(INSERT, backup_open_list_port[2][0])
        con4_on.insert(INSERT, backup_open_list_port[3][0])

        con1_off.insert(INSERT, backup_open_list_port[0][1])
        con2_off.insert(INSERT, backup_open_list_port[1][1])
        con3_off.insert(INSERT, backup_open_list_port[2][1])
        con4_off.insert(INSERT, backup_open_list_port[3][1])

def data_del_port(): #삭제하기
    global backup_list_port
    MsgBox = tk.messagebox.askquestion("삭제확인창", "일괄삭제 할까요?")
    if MsgBox == 'yes':
        speech_1_on_port = []
        backup_list_port = []
        backup_open_list_port = [["켜기 음성 입력하세요", "끄기 음성 입력하세요"],["켜기 음성 입력하세요", "끄기 음성 입력하세요"],["켜기 음성 입력하세요", "끄기 음성 입력하세요"],["켜기 음성 입력하세요", "끄기 음성 입력하세요"]]
        con1_on.delete(0, END)
        con2_on.delete(0, END)
        con3_on.delete(0, END)
        con4_on.delete(0, END)
        con1_off.delete(0, END)
        con2_off.delete(0, END)
        con3_off.delete(0, END)
        con4_off.delete(0, END)
        print("backup_open_list_port : ", backup_open_list_port)
        with open("back_up_port.txt", "w") as file:
            file.write("")
            file.write(f"{backup_open_list_port[0][0]}-{backup_open_list_port[0][1]}\n")
            file.write(f"{backup_open_list_port[1][0]}-{backup_open_list_port[1][1]}\n")
            file.write(f"{backup_open_list_port[2][0]}-{backup_open_list_port[2][1]}\n")
            file.write(f"{backup_open_list_port[3][0]}-{backup_open_list_port[3][1]}\n")

    else:
        tk.messagebox.showinfo('삭제확인창', '삭제가 취소되었습니다.')

    data_list_open_port()


def keyPressHandler(event):
    if (event.keycode==27):
        print(event.keycode)
        root.focus_set()

# bind the selected value changes
def focus_out(event):
    root.focus_set()

#_________

root = Tk()
root.title("음성 제어기 QVOICE_Ver6.0")
root.geometry("906x190+10+10")
root.resizable(False,False)

mic1_img = PhotoImage(file="mic1.png")
mic2_img = PhotoImage(file="mic2.png")
mic3_img = PhotoImage(file="mic3.png")
img_frm = Label(root, image=mic2_img)
img_frm.grid(column=0, row=0, rowspan=3,padx=4,pady=1)

file_entry = Text(root, width=52, height=0.5)
file_entry.grid(column=2, row=0, columnspan=5, pady=1)


port_label = Label(root, text="포트선택", font=("고딕", 10))
port_label.grid(column=0, row=3, pady=1)

label1 = Label(root, text="포트 선택", font=("고딕", 10))
label1.grid(column=7,columnspan=2, row=0)

label2 = Label(root, text="제어1", font=("고딕", 10))
label2.grid(column=7, row=2)
con1_on = Entry(root,width=20)
con1_on.grid(column=8, row=2, padx=2)
con1_off = Entry(root, width=20)
con1_off.grid(column=9, row=2,padx=2)

label3 = Label(root, text="제어2", font=("고딕", 10))
label3.grid(column=7, row=3)
con2_on = Entry(root, width=20)
con2_on.grid(column=8, row=3, padx=2)
con2_off = Entry(root, width=20)
con2_off.grid(column=9, row=3,padx=2)

label4 = Label(root, text="제어3", font=("고딕", 10))
label4.grid(column=7, row=4 )
con3_on = Entry(root, width=20)
con3_on.grid(column=8, row=4, padx=2)
con3_off = Entry(root, width=20)
con3_off.grid(column=9, row=4,padx=2)

label5 = Label(root, text="제어4", font=("고딕", 10))
label5.grid(column=7, row=5 )
con4_on = Entry(root, width=20)
con4_on.grid(column=8, row=5, padx=2)
con4_off = Entry(root, width=20)
con4_off.grid(column=9, row=5,padx=2)

lb = Listbox(root,width=92, height=4,bd=0 ,highlightthickness=0 ,font=("고딕",9))
lb.grid(column=0, row=3, rowspan=3,columnspan=7,padx=3,pady=4)

label6 = Label(root, text="선풍기", font=("고딕", 10))
label6.grid(column=7, row=6 )
lb6 = Label(root, text="[선풍기 켜, 선풍기 꺼, 선풍기 1단, 선풍기 2단,\n선풍기 3단, 선풍기 4단, 선풍기 회전]",foreground="dim gray", font=("고딕", 9))
lb6.grid(column=8, columnspan=2,rowspan=2,row=6 )

lb_content = Label(root, text="",foreground="dark goldenrod", font=("고딕", 9),wraplength=490)
lb_content.grid(column=0, columnspan=6,row=6 )

lb_count = Label(root, text="", font=("고딕", 10))
lb_count.grid(column=6, row=6 )

#참고result_label = Label(root, width=46, height=21, relief="ridge", wraplength=360, justify="left", font=("고딕", 10))

strs = StringVar()

Combx1 = ttk.Combobox(textvariable=strs, width=8) #콤보박스 선언
Combx1['value'] = (result) #콤보박스 요소 삽입
Combx1.bind('<<ComboboxSelected>>', focus_out)
Combx1.current(0) #0번째로 콤보박스

Combx1.grid(column=8,columnspan=2,row=0) #콤보박스 배치


speech_entry = Entry(root, width=20,)
speech_entry.grid(column=2, row=1, columnspan=2 ,padx=2,pady=1)

speech_off_entry = Entry(root, width=20,)
speech_off_entry.grid(column=5, row=1, columnspan=2,pady=1)

speech_off_label = Label(root, text="종료하기", font=("고딕", 10))
speech_off_label.grid(column=4, row=1, pady=1)

voice_label = Label(root, width=15,text="Voice_Cmd",foreground="blue",background="white",font=("고딕", 12, "bold"))
voice_label.grid(column=1, row=2,columnspan=2, pady=1)

#result_label = Label(root, width=51, height=9, relief="ridge", wraplength=400, justify="left", font=("고딕", 10))
#result_label.grid(column=2, row=2, rowspan=4,columnspan=5,padx=5,ipadx=12)

btn_fopen = Button(root, text="파일찾기", width=8, font=("고딕", 10), command=data_fopen)
btn_fopen.grid(column=1, row=0, pady=4)

btn_voice = Button(root, text="음성입력", state=tk.DISABLED, width=8, font=("고딕", 10),command=data_voice_save)
btn_voice.grid(column=1, row=1,  pady=1)

btn_list_del = Button(root, text="선택삭제", width=8, font=("고딕", 10), command=data_list_del)
btn_list_del.grid(column=4, row=2, pady=1)
btn_list_del.bind("<KeyPress>",data_del)

btn_del = Button(root, text="전부삭제", width=8, font=("고딕", 10), command=data_del)
btn_del.grid(column=5, row=2, pady=1)

btn_help = Button(root, text="도움말", width=8,font=("고딕", 10), command=help)
btn_help.grid(column=6, row=2, pady=1)

btn_open = Button(root, text="목록보기", width=8, font=("고딕", 10), command=data_list_open)
btn_open.grid(column=3, row=2,pady=1)

btn_on = Button(root, text="등 록", width=8, font=("고딕", 10), command=data_voice_save_port)
btn_on.grid(column=8, row=1,pady=1)

btn_off = Button(root, text="삭 제", width=8, font=("고딕", 10), command=data_del_port)
btn_off.grid(column=9, row=1,pady=1)


myThread = voiceloop()
myThread.rflag = True
myThread.start()

data_list_open() #초기함수
data_back() #초기함수
data_list_open() #초기함수
data_list_open_port()

root.bind("<KeyPress>",keyPressHandler)
root.protocol("WM_DELETE_WINDOW")
root.wm_attributes("-topmost", 1)
root.mainloop()