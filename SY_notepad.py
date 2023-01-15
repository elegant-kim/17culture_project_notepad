from tkinter import *
from tkinter.filedialog import *
from tkinter import font
from tkinter import colorchooser
from tkinter import messagebox
import os, sys
import win32print
import win32api
import PyPDF2

top = Tk()
# 타이틀 바에 프로그램명 표시
top.title("SY의 메모장!")
# 가로 1000 세로 500 크기의 메모장
top.geometry("1200x700")

# 파일명 변수를 글로벌로
global openstatusName
openstatusName = False
# 텍스트 선택영역변수를 글로벌로
global selected
selected = False


# 기능 1 : 새파일 생성
def newFile():
    try:
        # 새파일을 생성하면 Text Box 윈도우(myText)에 이전에 기록한 텍스트는 삭제
        myText.delete(1.0, END)
        # 타이틀 바에 파일명 표시
        top.title('새파일 - SY의 메모장!')
        # 상태 바에 새파일임을 표시
        statusBar.config(text="새파일      ")
        # 파일오픈 상태플래그는 False로
        global openstatusName
        openstatusName = False
    except Exception as ex:
        print('newFile() ' + str(ex))

# 기능 2 : 파일 불러오기 (텍스트 기반의 txt, html, py 등)
def openFile():
    try:
        # 이전 텍스트 삭제
        myText.delete(1.0, END)
        # 파일오픈 다이얼로그박스 불러오고, 선택된 파일명을 변수에 저장
        textFile = askopenfilename(title="Oepn File", filetypes=(("Text Files", "*.txt"), ("HTML Files", "*.html"),
                                                                 ("Python Files", "*.py"), ("All Files", "*.*")))
        # 파일명이 있는지 확인
        if textFile:
            # 나중에 액세스할 수 있도록 파일명을 글로벌로 만듬
            global openstatusName
            openstatusName = textFile
            # 상태 바에 파일명 표시
            name = textFile
            statusBar.config(text=f'{name}      ')
            # 타이틀 바에 파일명 표시
            top.title(f'{name} - SY의 메모장!')
            # 파일을 open하고,
            textFile = open(textFile, 'r', encoding='UTF8')
            # 파일의 내용을 읽어서 readContents 라는 변수에 임시 저장
            readContents = textFile.read()
            # Text Box 윈도우에 파일내용(readContents)을 출력(insert)
            myText.insert(END, readContents)
            # 파일 close
            textFile.close()
    except Exception as ex:
        print('openFile() ' + str(ex))

# 기능 3 : 다른이름으로 파일저장
def saveasFile():
    try:
        # 파일저장 다이얼로그박스 불러오고, 선택된 파일명을 변수에 저장
        textFile = asksaveasfilename(defaultextension=".txt", title="Save File",
                                     filetypes=(("Text Files", "*.txt"), ("HTML Files", "*.html"), ("Python Files", "*.py"),
                                                ("All Files", "*.*")))
        if textFile:
            # 상태 바에 파일명 표시
            name = textFile
            statusBar.config(text=f'저장: {name}      ')
            # 타이틀 바에 파일명 표시
            top.title(f'{name} - SY의 메모장!')
            # 파일을 open한 후 화면의 text내용을 파일에 저장
            textFile = open(textFile, 'w', encoding='UTF8')
            textFile.write(myText.get(1.0, END))
            # 파일 close
            textFile.close()
    except Exception as ex:
        print('saveasFile() ' + str(ex))

# 기능 3-1 : 파일 저장하기
def saveFile():
    try:
        global openstatusName
        if openstatusName:
            # 파일 저장
            textFile = open(openstatusName, 'w', encoding='UTF8')
            textFile.write(myText.get(1.0, END))
            # 파일 close
            textFile.close()
            # 상태바 업데이트
            statusBar.config(text=f'저장: {openstatusName}      ')
        else:
            saveasFile()
    except Exception as ex:
        print('saveFile() ' + str(ex))

# 기능 4 :  파일 출력 기능 (파일다이얼로그에서 파일선택하면 디폴트 설정된 프린터로 바로 출력)
def printFile():
    try:
        printerName = win32print.GetDefaultPrinter()
        statusBar.config(text=printerName)
        # 파일 다이얼로그박스 불러오고, 프린트할 파일변수에 저장
        filetoPrint = askopenfilename(title="Oepn File", filetypes=(("Text Files", "*.txt"), ("HTML Files", "*.html"),
                                                                     ("Python Files", "*.py"), ("All Files", "*.*")))
        if filetoPrint:
            win32api.ShellExecute(0, "print", filetoPrint, None, ".", 0)
    except Exception as ex:
        print('saveFile() ' + str(ex))


# 기능 5 : 텍스트 편집 기능 중 잘라내기 기능
def cutText(e):
    try:
        global selected
        # 키보드 shortcut을 사용하는지 체크
        if e:
            selected = top.clipboard_get()

        if myText.get(1.0, END) == '\n':
            pass
        else:
            if myText.selection_get():
                # 선택된 텍스트를 획득
                selected = myText.selection_get()
                # 선택된 영역 지우기
                myText.delete("sel.first", "sel.last")
                # 클립보드를 지우고 새로 append
                top.clipboard_clear()
                top.clipboard_append(selected)
    except Exception as ex:
        print('cutText() ' + str(ex))

# 기능 6 : 텍스트 편집 기능 중 복사 기능
def copyText(e):
    try:
        global selected
        # 키보드 shortcut을 사용하는지 체크
        if e:
            selected = top.clipboard_get()

        if myText.get(1.0, END) == '\n':
            pass
        else:
            if myText.selection_get():
                # 선택된 텍스트를 획득
                selected = myText.selection_get()
                # 클립보드를 지우고 새로 append
                top.clipboard_clear()
                top.clipboard_append(selected)
    except Exception as ex:
        print('copyText() ' + str(ex))

# 기능 7 : 텍스트 편집 기능 중 붙여넣기 기능
def pasteText(e):
    try:
        global selected
        # 키보드 shortcut을 사용하는지 체크
        if e:
            selected = top.clipboard_get()
        else:
            if selected:
                position = myText.index(INSERT)
                myText.insert(position, selected)
    except Exception as ex:
        print('pasteText() ' + str(ex))

# 기능 8 : 편집시 전체 선택 기능
def selectAll(e):
    try:
        myText.tag_add('sel', '1.0', 'end')
    except Exception as ex:
        print('selectAll() ' + str(ex))

# 기능 9 : 전체 선택된 내용 지우기
def clearAll():
    try:
        myText.delete(1.0, END)
    except Exception as ex:
        print('clearAll() ' + str(ex))

# 기능 10 : 폰트를 굵은글씨체로 변경하는 기능
def boldIt():
    try:
        if myText.get(1.0, END) == '\n':
            pass
        else:
            if myText.selection_get():
                # 폰트 생성
                boldFont = font.Font(myText, myText.cget("font"))
                boldFont.configure(weight="bold")
                # tag 구성
                myText.tag_configure("bold", font=boldFont)
                # 현재의 tag 지정
                current_tags = myText.tag_names("sel.first")
                # tag가 설정되었는지?
                if "bold" in current_tags:
                    myText.tag_remove("bold", "sel.first", "sel.last")
                else:
                    myText.tag_add("bold", "sel.first", "sel.last")
    except Exception as ex:
        print('boldIt() ' + str(ex))

# 기능 11 : 텍스트색상 변환 기능
def textColor():
    try:
        # 컬러를 선택
        myColor = colorchooser.askcolor()[1]
        if myColor:
            # 폰트 생성
            colorFont = font.Font(myText, myText.cget("font"))
            # tag 구성
            myText.tag_configure("colored", font=colorFont, foreground=myColor)
            # 현재의 tag 지정
            current_tags = myText.tag_names("sel.first")
            # tag가 설정되었는지?
            if "colored" in current_tags:
                myText.tag_remove("colored", "sel.first", "sel.last")
            else:
                myText.tag_add("colored", "sel.first", "sel.last")
    except Exception as ex:
        print('textColor() ' + str(ex))

# 기능 12 : 전체텍스트색상 변환 기능
def alltextColor():
    try:
        # 컬러를 선택
        myColor = colorchooser.askcolor()[1]
        if myColor:
            myText.config(fg=myColor)
    except Exception as ex:
        print('alltextColor() ' + str(ex))

# 기능 13 : 백그라운드 색상 변환 기능
def bgColor():
    try:
        # 컬러를 선택
        myColor = colorchooser.askcolor()[1]
        if myColor:
            myText.config(bg=myColor)
    except Exception as ex:
        print('bgColor() ' + str(ex))

# 기능 14 : 야간모드 변환 기능
def nightOn():
    try:
        # 전체색상을 정의
        bgmainColor = "#000000"
        bgsecondColor = "#424242"
        textmainColor = "green"
        textsecondColor = "#E0E0E0"
        # 전체 윈도우 색상 세팅
        top.config(bg=bgmainColor)
        statusBar.config(bg=bgmainColor, fg=textmainColor)
        myText.config(bg=bgsecondColor)
        toolbarFrame.config(bg=bgsecondColor)
        # 메뉴바 버튼 색상 세팅
        fileMenu.config(bg=bgsecondColor, fg=textsecondColor)
        editMenu.config(bg=bgsecondColor, fg=textsecondColor)
        optionMenu.config(bg=bgsecondColor, fg=textsecondColor)
        helpMenu.config(bg=bgsecondColor, fg=textsecondColor)
        # 툴바 버튼 색상 세팅
        newfileButton.config(bg=bgsecondColor, fg=textsecondColor)
        openfileButton.config(bg=bgsecondColor, fg=textsecondColor)
        savefileButton.config(bg=bgsecondColor, fg=textsecondColor)
        nightonButton.config(bg=bgsecondColor, fg=textsecondColor)
        nightoffButton.config(bg=bgsecondColor, fg=textsecondColor)
        boldButton.config(bg=bgsecondColor, fg=textsecondColor)
        textcolorButton.config(bg=bgsecondColor, fg=textsecondColor)
        openpdfButton.config(bg=bgsecondColor, fg=textsecondColor)
    except Exception as ex:
        print('nightOn() ' + str(ex))

# 기능 15 : 야간모드 끄기 기능
def nightOff():
    try:
        # 전체색상을 정의
        bgmainColor = "SystemButtonFace"
        bgsecondColor = "SystemButtonFace"
        bgthirdColor = "#795548"
        textmainColor = "black"
        textsecondColor = "SystemButtonFace"
        # 전체 윈도우 색상 세팅
        top.config(bg=bgmainColor)
        statusBar.config(bg=bgmainColor, fg=textmainColor)
        myText.config(bg="white")
        toolbarFrame.config(bg=bgsecondColor)
        # 메뉴바 버튼 색상 세팅
        fileMenu.config(bg=bgsecondColor, fg=textmainColor)
        editMenu.config(bg=bgsecondColor, fg=textmainColor)
        optionMenu.config(bg=bgsecondColor, fg=textmainColor)
        helpMenu.config(bg=bgsecondColor, fg=textmainColor)
        # 툴바 버튼 색상 세팅
        newfileButton.config(bg=bgthirdColor, fg=textsecondColor)
        openfileButton.config(bg=bgthirdColor, fg=textsecondColor)
        savefileButton.config(bg=bgthirdColor, fg=textsecondColor)
        nightonButton.config(bg=bgthirdColor, fg=textsecondColor)
        nightoffButton.config(bg=bgthirdColor, fg=textsecondColor)
        boldButton.config(bg=bgthirdColor, fg=textsecondColor)
        textcolorButton.config(bg=bgthirdColor, fg=textsecondColor)
        openpdfButton.config(bg=bgthirdColor, fg=textsecondColor)
    except Exception as ex:
        print('nightOff() ' + str(ex))

# 기능 16 : pdf 파일 열고 화면에 출력(텍스트 기반 PDF파일)
def openPDF():
    try:
        # 파일오픈 다이얼로그박스 불러오고, 선택된 PDF파일명을 변수에 저장
        openFile = askopenfilename(title="Oepn PDF File", filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")))
        # PDF파일을 선택했으면 읽음
        if openFile:
            # 상태 바에 pdf파일명 표시
            name = openFile
            statusBar.config(text=f'{name}      ')
            # 타이틀 바에 pdf파일명 표시
            top.title(f'{name} - SY의 메모장!')
            # PDF 파일을 open
            pdfFile = PyPDF2.PdfFileReader(openFile)
            # 읽을 page를 세팅
            page = pdfFile.getPage(0)
            # pdf 파일로부터 텍스트를 추출
            pageContents = page.extractText()
            # text Box 윈도우에 텍스트 추가
            myText.insert(1.0, pageContents)

    except Exception as ex:
        print('openPDF() ' + str(ex))

# 기능 17 : 프로그램 정보 디스믈레이
def aboutProgram():
    try:
        messagebox.showinfo("SY의 메모장!", "Version 0.01 (Build : 0.1.1.35)\n\n개발자 : SY\n\n"
                                         "파이썬으로 처음 만들어본 저의 메모장입니다.\n관심가져주셔서 감사합니다.")
    except Exception as ex:
        print('aboutProgram() ' + str(ex))

#### 스크롤바 적용을 위해 기존 소스 주석처리
# myText = Text(top)
# # 텍스트 입력부분 크기에 맞게 전체 윈도우 사이즈 조절
# top.grid_rowconfigure(0, weight=1)
# top.grid_columnconfigure(0, weight=1)
# # 텍스트가 4면을 모두 다 채우도록 고정
# myText.grid(sticky = N + E + S + W)
#### 스크롤바 적용을 위해 기존 소스 주석처리

# 툴바프레임 생성
toolbarFrame = Frame(top)
toolbarFrame.pack(fill=X)

# 메인프레임 생성
myFrame = Frame(top)
myFrame.pack(pady=5)
# 가로 스크롤바
horiScroll = Scrollbar(myFrame, orient='horizontal')
horiScroll.pack(side=BOTTOM, fill=X)
# 세로 스크롤바
vertScroll = Scrollbar(myFrame)
vertScroll.pack(side=RIGHT, fill=Y)
# (가로, 세로 스크롤바가 포함된) Text Box 윈도우 생성
myText = Text(myFrame, width=200, height=25, font=('맑은 고딕', 14), selectbackground="black", selectforeground="white",
              undo=True, xscrollcommand=horiScroll.set, yscrollcommand=vertScroll.set, wrap="none")
myText.pack()
# 스크롤바 구성
horiScroll.config(command=myText.xview)
vertScroll.config(command=myText.yview)

# Menu 생성
# file = None
myMenu = Menu(top)
top.config(menu=myMenu)
# File Menu 추가
fileMenu = Menu(myMenu, tearoff=False)
myMenu.add_cascade(label="파일", menu=fileMenu)
fileMenu.add_command(label="새파일", command=newFile)
fileMenu.add_command(label="열기", command=openFile)
fileMenu.add_command(label="저장", command=saveFile)
fileMenu.add_command(label="다른이름으로 저장", command=saveasFile)
fileMenu.add_separator()
fileMenu.add_command(label="파일 출력", command=printFile)
fileMenu.add_separator()
fileMenu.add_command(label="종료", command=top.quit)
# Edit Menu 추가
editMenu = Menu(myMenu, tearoff=False)
myMenu.add_cascade(label="편집", menu=editMenu)
editMenu.add_command(label="잘라내기", command=lambda: cutText(False), accelerator="Ctrl+X")
editMenu.add_command(label="복사", command=lambda: copyText(False), accelerator="Ctrl+C")
editMenu.add_command(label="붙여넣기        ", command=lambda: pasteText(False), accelerator="Ctrl+V")
editMenu.add_separator()
editMenu.add_command(label="Undo", command=myText.edit_undo, accelerator="Ctrl+Z")
editMenu.add_command(label="Redo", command=myText.edit_redo, accelerator="Ctrl+Y")
editMenu.add_separator()
editMenu.add_command(label="전체선택", command=lambda: selectAll(True), accelerator="Ctrl+A")
editMenu.add_command(label="지우기", command=clearAll)
# Option Menu 추가
optionMenu = Menu(myMenu, tearoff=False)
myMenu.add_cascade(label="부가기능", menu=optionMenu)
optionMenu.add_command(label="PDF파일열기", command=openPDF)
optionMenu.add_separator()
optionMenu.add_command(label="굵은글씨체", command=boldIt)
optionMenu.add_command(label="선택영역색상변환", command=textColor)
optionMenu.add_command(label="전체영역색상변환", command=alltextColor)
optionMenu.add_command(label="배경색상변환", command=bgColor)
optionMenu.add_separator()
optionMenu.add_command(label="야간모드", command=nightOn)
optionMenu.add_command(label="야간모드 끄기", command=nightOff)
# 도움말 Menu 추가
helpMenu = Menu(myMenu, tearoff=False)
myMenu.add_cascade(label="도움말", menu=helpMenu)
helpMenu.add_command(label="SY의 메모장 정보", command=aboutProgram)

# 상태바 추가
statusBar = Label(top, text='준비됨   ', anchor=E)
statusBar.pack(fill=X, side=BOTTOM, ipady=15)
# Edit 바인딩
top.bind('<Control-Key-x>', cutText)
top.bind('<Control-Key-c>', copyText)
top.bind('<Control-Key-v>', pasteText)
# Select 바인딩
top.bind('<Control-Key-a>', selectAll)

fee = "SongYeon"
myLabel = Label(top, text=fee[:-1]).pack()

# 메뉴 버튼 생성
# 새파일
newfileButton = Button(toolbarFrame, text="새파일", bg="#795548", fg="white", command=newFile)
newfileButton.grid(row=0, column=0, sticky=W, padx=5)
# 파일 오픈
openfileButton = Button(toolbarFrame, text="파일열기", bg="#795548", fg="white", command=openFile)
openfileButton.grid(row=0, column=1, padx=5)
# 파일 저장
savefileButton = Button(toolbarFrame, text="파일저장", bg="#795548", fg="white", command=saveFile)
savefileButton.grid(row=0, column=2, padx=5)
# 굵은글씨체 변경
boldButton = Button(toolbarFrame, text="굵은글씨", bg="#795548", fg="white", command=boldIt)
boldButton.grid(row=0, column=3, padx=5)
# Text Color
textcolorButton = Button(toolbarFrame, text="글씨색깔", bg="#795548", fg="white", command=textColor)
textcolorButton.grid(row=0, column=4, padx=5)
# night on 기능
nightonButton = Button(toolbarFrame, text="야간모드", bg="#795548", fg="white", command=nightOn)
nightonButton.grid(row=0, column=5, padx=5)
# night off 기능
nightoffButton = Button(toolbarFrame, text="야간모드 끄기", bg="#795548", fg="white", command=nightOff)
nightoffButton.grid(row=0, column=6, padx=5)
# pdf파일 읽기 기능
openpdfButton = Button(toolbarFrame, text="PDF파일열기", bg="#795548", fg="white", command=openPDF)
openpdfButton.grid(row=0, column=7, padx=5)

top.mainloop()