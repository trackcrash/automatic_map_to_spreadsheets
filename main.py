import gspread
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time

class AutoMammaDispenser:
    def __init__(self):
        self.root = tk.Tk("자동 맘마 디스펜서")
        self.root.geometry("300x300")
        self.chrome_options = webdriver.ChromeOptions()
        self.chrome_options.add_extension('extension_5_6_0_0.crx')
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
       
        #제목 필터링 단어
        self.filter_words = ['inst', 'instrument', 'cover', 'sfc', 'arrange', 'ピアノ','歌ってみた']
        tk.Label(self.root, text="구글 스프레드시트 링크").grid(row=0)
        tk.Label(self.root, text="시작 시간(초)").grid(row=1)
        tk.Label(self.root, text="곡수(시트번호:곡수,)").grid(row=2)

        self.e1 = tk.Entry(self.root)
        self.e2 = tk.Entry(self.root)
        self.e3 = tk.Entry(self.root)

        self.e1.grid(row=0, column=1)
        self.e2.grid(row=1, column=1)
        self.e3.grid(row=2, column=1)

        tk.Button(self.root, text='Submit', command=self.submit).grid(row=4, column=1, sticky=tk.W, pady=4)

    def run(self):
        self.root.mainloop()

    def submit(self):
        seen_title = set()
        spreadsheet_url = self.e1.get()
        start_time = int(self.e2.get())
        sheet_numbers_and_song_counts_str = self.e3.get()
        sheet_numbers_and_song_counts = [item.split(':') for item in sheet_numbers_and_song_counts_str.split(',')]
        self.cred = filedialog.askopenfilename(initialdir="./", title="Select file",
                                                  filetypes=(("json files", "*.json"), ("all files", "*.*")))
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(self.cred, self.scope)
        self.client = gspread.authorize(self.creds)
        doc = self.client.open_by_url(spreadsheet_url)

        result = []
        driver = webdriver.Chrome('./chromedriver.exe', options=self.chrome_options)

        for sheet_number_str, song_count_str in sheet_numbers_and_song_counts:
            sheet_number = int(sheet_number_str)
            song_count = int(song_count_str)
            
            sheet = doc.get_worksheet(sheet_number)
            data = sheet.get_all_records()
            df = pd.DataFrame(data)
            #sample = df.sample(song_count)
            sample = df.sample(frac=1).head(song_count)
            for index, row in sample.iterrows():
                anime_title = row['애니제목']
                song_title = row['곡제목']
                if anime_title == song_title:  # 만약 애니제목과 곡제목이 같다면
                    seen_title.add(anime_title)  # 해당 제목을 저장합니다.

                if anime_title in seen_title or song_title in seen_title:  # 만약 애니제목이나 곡제목 중 하나가 저장된 적이 있다면
                    continue
                try:
                    song_link = self.get_youtube_link(driver, row['애니 원제'] + ' ' + row['곡 원제'])
                    if song_link is None:
                        continue
                    result.append({
                        '정답': row['애니제목'],
                        '초성 힌트용 한글명칭': row['초성힌트용'],
                        '곡 제목': row['곡제목'],
                        '범주': 1,
                        '복수정답': row['애니판정'],
                        '음악 링크': song_link,
                        '미사용 여부': 'FALSE',
                        '시작/종료 시간': self.format_time(start_time),
                        '힌트 문구': row['사용처'],
                        '음악 판정': row['노래판정']
                    })
                except TimeoutException:
                    print(f"TimeoutException: {row['애니 원제']} {row['곡 원제']}")
                    continue
        driver.quit()
        result_df = pd.DataFrame(result)
        result_df.to_excel("output.xlsx", index=False)
        messagebox.showinfo("완료", "엑셀파일 생성완료.")

    def format_time(self, time_input):
        minutes = time_input // 60
        seconds = time_input % 60
        if time_input == 0:
            return ""
        return f"{minutes}:{seconds:02d}~"

    def get_youtube_link(self, driver, search_term):
        driver.get('https://www.youtube.com')

        wait = WebDriverWait(driver, 10)
        search_box = wait.until(EC.presence_of_element_located((By.NAME, 'search_query')))

        search_box.send_keys(search_term)
        search_box.send_keys(Keys.RETURN)

        videos = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div#container.style-scope.ytd-search a#thumbnail.yt-simple-endpoint.inline-block.style-scope.ytd-thumbnail')))
        durations = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                                  'div#container.style-scope.ytd-search span#text.style-scope.ytd-thumbnail-overlay-time-status-renderer')))

        for video, duration in zip(videos, durations):
            video_url = video.get_attribute('href')
            if video_url is None:
                continue
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
                if total_seconds < 80 or total_seconds > 420:
                    print(f"넘길엉:{search_term}")
                    continue
                if self.filtering(video_title):
                    print(f"커버시러퉤퉤퉷:{search_term}")
                    continue
            return clean_url
        return None

    def filtering(self, video_title):
        return any(word in video_title.lower() for word in self.filter_words)


if __name__ == '__main__':
    dispenser = AutoMammaDispenser()
    dispenser.run()
