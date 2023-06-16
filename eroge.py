import gspread
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time

#TK
root = tk.Tk("자동 맘마 디스펜서")
root.geometry("300x300")
# chrome_options = Options()
# chrome_options.add_argument("--headless")
chrome_options = webdriver.ChromeOptions()
chrome_options.add_extension('extension_5_6_0_0.crx')
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
#cred = tk.filedialog.askopenfilename(initialdir = "./",title = "Select file",filetypes = (("json files","*.json"),("all files","*.*")))
#스프레드 시트 api넣기 나중에 파일선택으로 교체해주기
creds = ServiceAccountCredentials.from_json_keyfile_name('api-project-92887753-a2f2c523fcd2.json', scope)
client = gspread.authorize(creds)
#비디오 길이 체크
def check_video_lengths():
    df = pd.read_excel('output.xlsx')
    driver = webdriver.Chrome('./chromedriver.exe', options=chrome_options)
    time.sleep(20)
    valid_videos = []

    for index, row in df.iterrows():
        video_url = row['음악 링크']
        driver.get(video_url)
        
        try:
            duration_str = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span.ytp-time-duration'))).text
        except TimeoutException:
            print(f"Could not load video length for {video_url}")
            continue

        time_parts = duration_str.split(':')
        if len(time_parts) == 3:
            hours, minutes, seconds = map(int, time_parts)
            total_seconds = hours * 3600 + minutes * 60 + seconds
        elif len(time_parts) == 2:
            minutes, seconds = map(int, time_parts)
            total_seconds = minutes * 60 + seconds
        else:
            continue

        if 80 <= total_seconds <= 420:
            valid_videos.append(row)
        else:
            print(f"Video length out of range for {video_url}: {duration_str}")

    driver.quit()

    valid_videos_df = pd.DataFrame(valid_videos)
    valid_videos_df.to_excel("output.xlsx", index=False)
#시간 변환
def format_time(time_input):
    minutes = time_input // 60
    seconds = time_input % 60
    if time_input == 0:
        #아무것도 리턴하지 않음
        return ""
    return f"{minutes}:{seconds:02d}~"
#유튜브에서 검색후 첫번째요소 리턴
def get_youtube_link(driver, search_term):
    driver.get('https://www.youtube.com')
    
    # 'search_query' 요소가 로드될 때까지 최대 10초 동안 기다립니다.
    wait = WebDriverWait(driver, 10)
    search_box = wait.until(EC.presence_of_element_located((By.NAME, 'search_query')))

    search_box.send_keys(search_term)
    search_box.send_keys(Keys.RETURN)

    # 'video-title' 요소가 로드될 때까지 최대 10초 동안 기다립니다.
    videos = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,'div#container.style-scope.ytd-search a#thumbnail.yt-simple-endpoint.inline-block.style-scope.ytd-thumbnail')))
    durations = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div#container.style-scope.ytd-search span#text.style-scope.ytd-thumbnail-overlay-time-status-renderer')))

    for video, duration in zip(videos, durations):
        video_url = video.get_attribute('href')
        video_title = video.get_attribute('title')
        clean_url = video_url.split("&")[0]
        duration_str = duration.get_attribute('innerText').strip()
        time_parts = duration_str.split(':')
        if 'youtube.com/shorts/' in clean_url:
            continue
        if 'youtube.com/user/' in clean_url:
            continue
        if len(time_parts) == 3:
            continue
        elif len(time_parts) == 2:
            minutes, seconds = map(int, time_parts)
            total_seconds = minutes * 60 + seconds
            print(minutes, seconds, total_seconds)
            if total_seconds < 80 or total_seconds > 420:
                print(f"넘길엉:{search_term}")
                continue
            if any(word in video_title.lower() for word in ['inst', 'instrument', 'cover', 'sfc','arrange','ピアノ','Arrange']): 
                print(f"커버시러퉤퉤퉷:{search_term}")
                continue
        return clean_url
    return None

#버튼 눌렀을 때 실행
def submit():
    spreadsheet_url = e1.get()
    number_of_songs = int(e2.get())
    start_time = int(e3.get())
    doc = client.open_by_url(spreadsheet_url)
    sheet = doc.get_worksheet(2)
    
    data = sheet.get_all_records()

    df = pd.DataFrame(data)
    
    sample = df.sample(number_of_songs)
    
    result = []
    driver = webdriver.Chrome('./chromedriver.exe')
    for index, row in sample.iterrows():
        try:
            song_link = get_youtube_link(driver,row['작품 원제'] + ' ' + row['곡 원제'])
            if song_link is None:  # 링크를 받아오지 못했으면 건너뛰기
                continue
            result.append({
                '정답': row['작품제목'],
                '초성 힌트용 한글명칭': row['초성힌트용'],
                '곡 제목': row['곡제목'],
                '범주': 1,
                '복수정답': row['작품판정'],
                '음악 링크': song_link,
                '미사용 여부': 'FALSE',
                '시작/종료 시간': format_time(start_time),
                '힌트 문구': row['회사이름'],
                '음악 판정': row['노래판정']
            })
        except TimeoutException:
            print(f"TimeoutException: {row['작품 원제']} {row['곡 원제']}")
            continue
    driver.quit()
    result_df = pd.DataFrame(result)
    result_df.to_excel("eroge_output.xlsx", index=False)
    messagebox.showinfo("완료", "엑셀파일 생성완료.")

# setup UI
tk.Label(root, text="구글 스프레드시트 링크").grid(row=0)
tk.Label(root, text="곡수").grid(row=1)
tk.Label(root, text="시작 시간(초)").grid(row=2)

e1 = tk.Entry(root)
e2 = tk.Entry(root)
e3 = tk.Entry(root)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2, column=1)
#오른쪽 아래로 붙이기
tk.Button(root, text='Submit', command=submit).grid(row=3, column=1, sticky=tk.W, pady=4)

root.mainloop()
