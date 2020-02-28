#구글 로그인 상태여야 정상 작동합니다.
#실행은 키보드 말고 "Run" 버튼으로 실행하여 주세요!!
#작동이 안되면 첨부된 사진처럼 구글 편지쓰기 캡처를 다시해주세요.


import pyautogui as pag
import pyperclip
import time
import pandas as pd


def openGmail():
    pag.hotkey("winleft")
    time.sleep(1.0)
    pag.typewrite('chrome')
    time.sleep(1.5)
    pag.hotkey("enter")
    time.sleep(3.0)
    pag.typewrite('https://www.google.com/gmail/')
    time.sleep(1.5)
    pag.hotkey("enter")

    time.sleep(1.0)
    pag.hotkey("winleft", "up")
    time.sleep(5.0)
    return 0


def sendMail(emailAddress, title, contents):
    time.sleep(2.0)
    locate = pag.locateCenterOnScreen("create_32dp.png")
    # print(locate)
    pag.click(locate)
    time.sleep(1.0)

    pag.typewrite(emailAddress)  # 이메일 주소
    time.sleep(0.5)
    pag.hotkey("enter")
    time.sleep(0.5)
    pag.hotkey("tab")
    time.sleep(0.5)

    pyperclip.copy(title)  # 제목
    pag.hotkey("ctrl", "v")
    time.sleep(0.5)
    pag.hotkey("tab")
    time.sleep(0.5)

    pyperclip.copy(contents)  # 메일 내용
    pag.hotkey("ctrl", "v")
    time.sleep(0.5)

    pag.hotkey("tab")
    time.sleep(0.5)
    pag.hotkey("enter")
    time.sleep(0.5)
    return 0


def entranceMoney(devision):
    if (
            devision == "기계공학과" or devision == "산업공학과" or devision == "화학공학과" or devision == "신소재공학과" or devision == "응용화학생명공학과" or devision == "환경안전공학과" or devision == "건설시스템공학과" or devision == "교통시스템공학과" or devision == "건축학과" or devision == "융합시스템공학과"):
        money = "4,894,000"
    elif (devision == "전자공학과" or devision == "소프트웨어학과" or devision == "사이버보안학과" or devision == "미디어학과"):
        money = "4,894,000"
    elif (devision == "국방디지털융합학과"):
        money = "5,309,000"
    elif (devision == "수학과" or devision == "물리학과" or devision == "화학과" or devision == "생명과학과"):
        money = "4,384,000"
    elif (devision == "의학과"):
        money = "5,578,000"
    elif (devision == "간호학과"):
        money = "4,384,000"
    elif (devision == "경영학과" or devision == "글로벌경영학과"):
        money = "3,908,000"
    elif (devision == "e-비즈니스학과"):
        money = "4,345,000"
    elif (devision == "금융공학과"):
        money = "4,576,000"
    elif (
            devision == "국어국문학과" or devision == "영어영문학과" or devision == "영어영문학과" or devision == "불어불문학과" or devision == "사학과" or devision == "문화콘텐츠학과" or devision == "경제학과" or devision == "행정학과" or devision == "심리학과" or devision == "사회학과" or devision == "정치외교학과" or devision == "스포츠레저학과"):
        money = "3,842,000"

    return money


personList = pd.read_excel("info.xlsx")
personList = personList.values.tolist()
# print(personList)

openGmail()
for i in range(len(personList)):
    emailAddress = personList[i][4]
    if (personList[i][2] == 20):
        title = "아주대 합격을 축하합니다. 등록금 및 기숙사 관련 안내사항입니다."
    else:
        title = "아주대 새 학기 등록금 및 생활관 관련 안내사항입니다."
    money = entranceMoney(personList[i][1])

    if (personList[i][5] == "남자"):
        liveRoom = "국제학사, 용지관, 남제관"
    else:
        liveRoom = "국제학사, 광교관"

    if (personList[i][3] == "강원도" or personList[i][3] == "부산광역시" or personList[i][3] == "광주광역시" or personList[i][
        3] == "대전광역시" or personList[i][3] == "경상북도"):
        score = 30
    elif (personList[i][3] == "서울특별시(북부)"):
        score = 15

    contents = "안녕하세요 " + str(personList[i][0]) + "님, 아주대 입학처입니다. " + str(personList[i][0]) + "님은 " + str(
        personList[i][1]) + "로 납부해야 하는 등록금은" + str(money) + "원 입니다. 생활관은 " + str(liveRoom) + "중에 입주 가능하고 " + str(
        personList[i][0]) + "님의 거리점수는 " + str(score) + "입니다. 더 자세한 사항은 http://iajou.uway.com/intro/intro.htm 에서 확인해주세요."
    sendMail(emailAddress, title, contents)
    pag.moveTo(925, 503, duration=0.1)

