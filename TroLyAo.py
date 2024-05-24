
import os

from gtts import gTTS
import playsound
import speech_recognition

from time import strftime
import time
import datetime
#Chọn ngẫu nhiên
import random
#Truy cập web, trình duyệt
import re
import webbrowser
#Lấy thông tin từ web
import requests
from datetime import *
import os
import speech_recognition as sr
from playsound import playsound
from gtts import gTTS
import smtplib
import xlrd
import sys
import wikipedia 
from qrcode import *
import json
#Truy cập web, trình duyệt, hỗ trợ tìm kiếm

from webdriver_manager.chrome import ChromeDriverManager
from youtube_search import YoutubeSearch
from youtubesearchpython import SearchVideos
#Thư viện Tkinter hỗ trợ giao diện
from tkinter import Tk, RIGHT, BOTH, RAISED
from tkinter.ttk import Frame, Button, Style
from tkinter import *
from PIL import Image, ImageTk
import tkinter.messagebox as mbox


root = Tk()
text_area = Text(root, height=26, width=45)
scroll = Scrollbar(root, command=text_area.yview)

def r_set():
    image_1 = ImageTk.PhotoImage(Image.open("image\\trolyao.jpg"))    
    label1 = Label(image=image_1)
    label1.image = image_1
    label1.place(x=7, y=43)
    text_area.delete("1.0", "1000000000.0")

def day():  
    ai_brain = "Hôm nay là ngày " + date.today().strftime("%d") + " Tháng " + date.today().strftime("%m") + " năm " + date.today().strftime("%Y")
    print(ai_brain)
    say(ai_brain)
def time():
    ai_brain = "Bây giờ là " + datetime.now().strftime("%H") + ' giờ ' + datetime.now().strftime("%M") + ' phút'    
    print(ai_brain)
    say(ai_brain)
def say(text):
    tts = gTTS(text=text,tld="com.vn",lang='vi',slow=False)
    date_string = datetime.now().strftime("%d%m%Y%H%M%S")
    file = "voice\\sound"+date_string+".mp3"
    tts.save(file)
    text_area.insert(INSERT,"IRIS: "+text+"\n")
    root.update()
    playsound(file)
    os.remove(file)


def listen():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Tôi: ", end='')
        audio = r.listen(source, phrase_time_limit=3)
        try:
            text = r.recognize_google(audio, language="vi-VN").lower()
            print(text)
            text_area.insert(INSERT,"Học sinh: "+text+"\n")
            root.update()
            return text
        except:
            say('Tôi không nghe thấy')
            return 'Tôi không nghe thấy'
def sendemail():
    #In, nói và nghe nhận biến nguoinhan
    ai_brain = 'Bạn gửi email cho ai nhỉ'
    print(ai_brain)
    say(ai_brain)
    nguoinhan = listen()
    print(nguoinhan)
    #từ biến nguoinhan ra email
    if 'h' in nguoinhan: email = 'hoangphamconglc2212@gmail.com'
    elif 'hưng' in nguoinhan: email = 'vingochung@gmail.com'
    elif 'hương' in nguoinhan: email = 'hahuong77@gmail.com'
    elif "khảo" in nguoinhan: email = 'luongvietha01@gmail.com'
    elif "hoàng" in nguoinhan: email = 'dangnhailc@gmail.com'
    #Gửi email
    if '@' in email:
        ai_brain = 'Nội dung bạn muốn gửi là gì'
        print(ai_brain)
        say(ai_brain)
        noidung = "[Lớp 10 Chuyên Toán] " + listen()
        print(noidung)
        mail = smtplib.SMTP('smtp.gmail.com', 587)
        mail.ehlo()
        mail.starttls()
        mail.login('duankhoahockithuat@gmail.com', 'hoanglc2212')
        mail.sendmail('duankhoahockithuat@gmail.com',email, noidung.encode('utf-8'))
        mail.close()
        ai_brain = 'Email của bạn vừa được gửi'
        say(ai_brain)
        print(ai_brain)
    else:
        ai_brain = 'Người bạn cần gửi không có trong danh bạ'
def openchrome():
    ai_brain = "Mở Google Chrome"
    os.startfile('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')
    print(ai_brain)
    say(ai_brain)
def hello():
    day_time = int(datetime.now().strftime("%H"))
    if day_time < 12 :
        ai_brain = "Chào buổi sáng, chúc bạn một ngày tốt lành"
    elif 12 <= day_time < 18:
        ai_brain = "Xin chào, chúc một buổi chiều tốt lành"
    else:
        ai_brain = "Xin chào, chúc một buổi tối vui vẻ"
    print(ai_brain)
    say(ai_brain)
def bye():
    ai_brain = "Hẹn gặp lại"
    print(ai_brain)
    say(ai_brain)
    sys.exit()
def helpme():
    say("""Tôi có thể giúp bạn thực hiện các câu lệnh sau đây:
    1. Chào hỏi
    2. Hiển thị giờ
    3. Xem điểm 
    4. Tìm kiếm trên Google
    5. Wikipedia
    6. Gửi email cho giáo viên
    7. Ký sổ đầu bài
     """)
def checkscore():
    ai_brain = "Bạn muốn xem điểm tổng kết hay điểm chi tiết?"
    say(ai_brain)
    print(ai_brain)
    ai_brain = listen()
    if "chi tiết" in ai_brain:    
        ai_brain = "Bạn muốn xem điểm của ai nhỉ? - Đọc họ và tên"
        say(ai_brain)
        print(ai_brain)
        nguoidiem = listen()
        print(nguoidiem)
        if nguoidiem == 'đặng quốc dương': file_location = 'D:\\Summer_project\\KHOA HỌC KĨ THUẬT\\_Testfiles\\Bảng điểm lớp\\Đặng Quốc Dương.xls'
        elif nguoidiem == 'nguyễn nam khánh': file_location = 'D:\\Summer_project\\KHOA HỌC KĨ THUẬT\\_Testfiles\\Bảng điểm lớp\\Nguyễn Nam Khánh.xls'
        elif nguoidiem == 'phạm bá huy': file_location = 'D:\\Summer_project\\KHOA HỌC KĨ THUẬT\\_Testfiles\\Bảng điểm lớp\\Phạm Bá Huy.xls'
        elif nguoidiem == 'phạm công hoàng': file_location = 'D:\\Summer_project\\KHOA HỌC KĨ THUẬT\\_Testfiles\\Bảng điểm lớp\\Phạm Công Hoàng.xls'
        elif nguoidiem =='phạm hoài nam': file_location = 'D:\\Summer_project\\KHOA HỌC KĨ THUẬT\\_Testfiles\\Bảng điểm lớp\\Phạm Hoài Nam.xls'
        else: 
            file_location = 'Không tìm thấy'
            print(file_location)
            say(file_location)
        if file_location != 'Không tìm thấy':
            wb = xlrd.open_workbook(file_location)
            sheet = wb.sheet_by_index(0)
            ai_brain = "Bạn muốn xem điểm môn gì"
            tuychon = """
                CÁC TÙY CHỌN
        + Toán
        + Văn
        + Anh
        + Toàn bộ
            """
            print(tuychon)
            say(tuychon)
            monhoc = listen()
            if 'Toán' in monhoc: 
                diem ='Hệ Số 1: ' + sheet.cell_value(4, 2) + sheet.cell_value(4, 3)  + '\nHệ số 2: ' +  sheet.cell_value(4, 12) + '\nHệ số 3: ' + sheet.cell_value(4, 19) + '\nTrung bình: ' + sheet.cell_value(4, 20)
                print(diem)
            elif 'Văn' in monhoc:
                diem ='Hệ Số 1: ' + sheet.cell_value(5, 2) + sheet.cell_value(5, 3)  + '\nHệ số 2: ' +  sheet.cell_value(5, 12) + '\nHệ số 3: ' + sheet.cell_value(5, 19) + '\nTrung bình: ' + sheet.cell_value(5, 20)
                print(diem)
            elif 'Anh' in monhoc:
                diem ='Hệ Số 1: ' + sheet.cell_value(6, 2) + sheet.cell_value(6, 3)  + '\nHệ số 2: ' +  sheet.cell_value(6, 12) + '\nHệ số 3: ' + sheet.cell_value(6, 19) + '\nTrung bình: ' + sheet.cell_value(6, 20)
            elif 'toàn bộ' in monhoc:
                ai_brain = 'Mở bảng điểm'
                print(ai_brain)
                say(ai_brain)
                os.startfile(file_location)
    if "tổng kết" in ai_brain:
        ai_brain = "Bạn muốn xem điểm của ai nhỉ? - Đọc họ và tên"
        say(ai_brain)
        print(ai_brain)
        nguoidiem = listen()        
        if nguoidiem == "lê thế anh" : stt = 7
        elif nguoidiem == "vũ quỳnh anh": stt = 8
        elif nguoidiem == "lê quang dũng" : stt = 9
        elif nguoidiem == "lê xuân dũng" : stt = 10
        elif nguoidiem == "đặng quốc dương" : stt = 11
        elif nguoidiem == "lê ngọc khánh đan" : stt = 12
        elif nguoidiem == "dương minh đức" : stt = 13
        elif nguoidiem == "nguyễn phương giang" : stt = 14
        elif nguoidiem == "trịnh hoàng đại hải" : stt = 15
        elif nguoidiem == "phạm công hoàng" : stt = 16
        elif nguoidiem == "phạm bá huy" : stt = 17
        elif nguoidiem == "phạm gia huy" : stt = 18
        elif nguoidiem == "nguyễn trọng hưng" : stt = 19
        elif nguoidiem == "nguyễn gia khánh" : stt = 20
        elif nguoidiem == "nguyễn nam khánh" : stt = 21
        elif nguoidiem == "trần nhật khánh" : stt = 22
        elif nguoidiem == "nguyễn trung kiên" : stt = 23
        elif nguoidiem == "hoàng khánh linh" : stt = 24
        elif nguoidiem == "vương phượng linh" : stt = 25
        elif nguoidiem == "nguyễn thành luân" : stt = 26
        elif nguoidiem == "đoàn duy mạnh" : stt = 27
        elif nguoidiem == "vương thành nam" : stt = 28
        elif nguoidiem == "nguyễn hà tuấn nghĩa" : stt = 29
        elif nguoidiem == "nguyễn yến ngọc" : stt = 30
        elif nguoidiem == "trần bích ngọc" : stt = 31
        elif nguoidiem == "hoàng đặng hà nhi" : stt = 32
        elif nguoidiem == "nguyễn yến nhi" : stt = 33
        elif nguoidiem == "trần thị hồng nhung" : stt = 34
        elif nguoidiem == "nguyễn hồng phương" : stt = 35
        elif nguoidiem == "thái đức thành" : stt = 36
        elif nguoidiem == "phạm minh thuận" : stt = 37
        elif nguoidiem == "đinh thị hoài thương" : stt = 38
        elif nguoidiem == "nguyễn thị cẩm tiên" : stt = 39
        elif nguoidiem == "hoàng thị kiều trang" : stt = 40
        elif nguoidiem == "nguyễn đặng anh tuấn" : stt = 41 
        else : 
            say("Không tìm thấy học sinh, xin mời tìm kiếm lại")
            return
        file_location = "C:\\Users\\Hoàng\\Desktop\\IRIS\\SK HKI 10 Toán.xls"
        wb = xlrd.open_workbook('C:\\Users\\Hoàng\\Desktop\\IRIS\\SK HKI 10 Toán.xls')
        sheet = wb.sheet_by_index(0)
        trungbinhmon = str(sheet.cell_value(stt, 2 )) +  "\nToán: " +  str(sheet.cell_value(stt, 4 )) + "\nLí: "+  str(sheet.cell_value(stt, 5))+ "\nHóa: "+  str(sheet.cell_value(stt, 6))+ "\nSinh: "+  str(sheet.cell_value(stt, 7))+ "\nTin: "+  str(sheet.cell_value(stt, 8))+ "\nVăn: "+  str(sheet.cell_value(stt, 9))+ "\nSử: "+  str(sheet.cell_value(stt, 10))+ "\nĐịa: "+  str(sheet.cell_value(stt, 11))+ "\nNgoại ngữ: "+  str(sheet.cell_value(stt, 12))+ "\nGiáo dục công dân: "+  str(sheet.cell_value(stt, 13))+ "\nCông nghệ: "+  str(sheet.cell_value(stt, 14))+ "\nGiáo dục quốc phòng: "+  str(sheet.cell_value(stt, 16))+ "\nTrung bình: "+  str(sheet.cell_value(stt, 17)) + "\nHọc lực: "+  str(sheet.cell_value(stt, 18)) +"\nHạnh kiểm: "+  str(sheet.cell_value(stt, 19)) 
        say(trungbinhmon)

def thoitiet():
    say("Bạn muốn xem thời tiết ở đâu ạ?")
    city = listen()
    ow_url = "http://api.openweathermap.org/data/2.5/weather?"
    if not city:
        pass
    api_key = "fe8d8c65cf345889139d8e545f57819a"
    call_url = ow_url + "appid=" + api_key + "&q=" + city + "&units=metric"
    response = requests.get(call_url)
    data = response.json()
    if data["cod"] != "404":
        city_res = data["main"]
        current_temperature = city_res["temp"]
        current_pressure = city_res["pressure"]
        current_humidity = city_res["humidity"]
        suntime = data["sys"]
        wthr = data["weather"]
        weather_description = wthr[0]["description"]
        content = """
Nhiệt độ trung bình là {temp} độ C
Áp suất không khí là {pressure} héc tơ Pascal
Độ ẩm là {humidity}%
        """.format(temp = current_temperature, pressure = current_pressure, humidity = current_humidity)
        say(content)
    else:
        say("Không tìm thấy địa chỉ của bạn")    
def wiki():
    try:
        wikipedia.set_lang('vi')
        say("Bạn muốn nghe về gì ạ")
        text = listen()
        contents = wikipedia.summary(text).split('\n')
        say(contents[0])
        print(contents)
        time.sleep(10)
        for content in contents[1:]:
            print("Bạn muốn nghe thêm không")
            say("Bạn muốn nghe thêm không")         
            ans = listen()
            if "có" not in ans:
                break    
            say(content)
            time.sleep(10)
        say('Cảm ơn bạn đã lắng nghe!!!')
    except:
        say("Tôi không định nghĩa được thuật ngữ của bạn. Xin mời bạn nói lại")
def qr_code():
        os.startfile('qrcode.png')

print('Trợ lí ảo - Made by KoNgWoAnG\n')

def ham_main():
    r = speech_recognition.Recognizer()
    you=""
    ai_brain=""
    while True:
        with speech_recognition.Microphone() as source:
            playsound("Ping.mp3", False)
            print("AI:  Dang nghe ...")
            audio = r.listen(source, phrase_time_limit=6)
            print("AI:  ...")
        try:
            you = r.recognize_google(audio, language="vi-VN")
            print("\nYou: "+ you)	
            you = you.lower()
        except:
            ai_brain = "Tôi nghe không rõ. Bạn nói lại được không"
            print("\nAI: " + ai_brain)

        text_area.insert(INSERT,"You: "+you+"\n")
        root.update()

        say("Bạn cần tôi hỗ trợ gì ạ?")
        print("Bạn cần tôi hỗ trợ gì ạ?")
        ai_ear = listen()
        if 'chào' in ai_ear or 'hello' in ai_ear: 
            hello()
        elif 'giúp' in ai_ear:
            helpme()
        elif 'tạm biệt' in ai_ear or 'thoát' in ai_ear:
            bye()
        elif 'giờ' in ai_ear:
            time()
        elif 'ngày' in ai_ear:
            day()
        elif 'chrome' in ai_ear or 'trình duyệt' in ai_ear:
            openchrome()
        elif 'thư' in ai_ear or 'email' in ai_ear or 'liên hệ' in ai_ear:
            sendemail()
        elif 'điểm' in ai_ear:
            checkscore()
        elif 'wikipedia' in ai_ear or 'thông tin' in ai_ear:
            wiki()
        elif 'đầu bài' in ai_ear:
            qr_code()
        else: 
            ai_brain = 'Tôi không thể giúp gì cho bạn. Bạn có thể tìm hỗ trợ khác'
            say(ai_brain)
            print(ai_brain)

        text_area.insert(INSERT,"_____________________________________________")
        root.update()
        you=""

class Example(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent) 
        self.parent = parent
        self.initUI()
    
    def initUI(self):
        self.parent.title("TRỢ LÝ ẢO - IRIS")
        self.style = Style()
        self.style.theme_use("default")
        
        scroll.pack(side=RIGHT, fill=Y)
        text_area.configure(yscrollcommand=scroll.set)
        text_area.pack(side=RIGHT)


        frame = Frame(self, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        self.pack(fill=BOTH, expand=True)

        closeButton = Button(self, text="Thoát",command = bye,width=10,fg="white", bg="#009999",bd=3)
        closeButton.pack(side=RIGHT, padx=11, pady=10)
        okButton = Button(self, text="Bắt đầu",command = ham_main,width=10,fg="white", bg="#009999",bd=3)
        okButton.pack(side=RIGHT, padx=11, pady=10)
        help1 = Button(self, text="Trợ Giúp",command = helpme,width=10,fg="white", bg="#009999",bd=3)
        help1.pack(side=RIGHT, padx=11, pady=10)
        lammoi = Button(self,text="Wikipedia",command = wiki,width=10,fg="white", bg="#009999",bd=3)
        lammoi.pack(side=RIGHT,padx=11, pady=10)
        trogiup = Button(self,text="Thời tiết",command = thoitiet,width=10,fg="white", bg="#009999",bd=3)
        trogiup.pack(side=RIGHT,padx=11, pady=10)
        openapp = Button(self,text="Mở Chrome",command = openchrome,width=10,fg="white", bg="#009999",bd=3)
        openapp.pack(side=RIGHT,padx=11, pady=10)
        openapp1 = Button(self,text="Xem điểm",command = checkscore,width=10,fg="white", bg="#009999",bd=3)
        openapp1.pack(side=RIGHT,padx=11, pady=10)
        openapp2 = Button(self,text="Liên hệ",command = sendemail,width=10,fg="white", bg="#009999",bd=3)
        openapp2.pack(side=RIGHT,padx=11, pady=10)
        openapp3 = Button(self,text="Sổ đầu bài",command = qr_code,width=10,fg="white", bg="#009999",bd=3)
        openapp3.pack(side=RIGHT,padx=11, pady=10)
        # self.pack(fill=BOTH, expand=1)   
        # Style().configure("TFrame", background="#333")
    
        image_1 = ImageTk.PhotoImage(Image.open("image\\trolyao.jpg"))    
        label1 = Label(self, image=image_1)
        label1.image = image_1
        label1.place(x=7, y=43)

        l = Label(root, text='BOX CHAT', fg='White', bg='blue')
        l.place(x = 1100, y = 10, width=120, height=25)
        l1 = Label(root, text='TRỢ LÝ ẢO IRIS', fg='yellow', bg='black')
        l1.place(x = 250, y = 11, width=120, height=25)

root.geometry("1600x510+250+50")
root.resizable(True, True)
app = Example(root)
root.mainloop()

#http://api.openweathermap.org/data/2.5/weather?appid=b0d4f9bfd2bbc40d10976e6fd3ea7514&q=da%20nang&units=metric