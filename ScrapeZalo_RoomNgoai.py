import time
import os
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from pathlib import Path
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta

def scrape_zalo_profile(ds_tengrp, ds_group_id, room_names):
    try:
        print("Initiating scrape for Zalo profiles")

        # Đường dẫn đến profile của Chrome
        service = Service("C:/Users/Welcome/Documents/Python/Scrape tiktok video/chromedriver-win64/chromedriver.exe")
        user_data_dir = Path("C:/Users/Welcome/AppData/Local/Google/Chrome/User Data")

        # Cấu hình ChromeOptions
        options = webdriver.ChromeOptions()
        # options.add_argument('headless')
        options.add_argument("disable-extensions")
        options.add_argument("--remote-debugging-port=9222") 
        options.add_argument("--user-data-dir={}".format(user_data_dir))
        options.add_argument('--profile-directory=Profile 13')

        # Mở trình duyệt
        browser = webdriver.Chrome(service=service, options=options)
        print("Browser started successfully")

        # Mở trang Zalo
        browser.get("https://chat.zalo.me/")
        print("Zalo profile page loaded successfully")
        time.sleep(6)  # Chờ trang tải

        # Tạo DataFrame tổng hợp
        df_full = pd.DataFrame()

        # Lặp qua danh sách group chat
        for tengrp, group_id, room_name in zip(ds_tengrp, ds_group_id, room_names):
            print(f"Đang xử lý group chat: {tengrp}")

            # Bấm vào ô tìm kiếm và nhập tên group chat
            browser.find_element(By.XPATH, '//input[@id="contact-search-input"]').click()
            time.sleep(2)
            browser.find_element(By.XPATH, '//input[@id="contact-search-input"]').send_keys(tengrp)
            time.sleep(2)

            # Click vào group chat tìm thấy
            browser.find_element(By.XPATH, f'//div[@id="{group_id}"]').click()
            time.sleep(2)

            # Cuộn và kiểm tra
            while True:
                scroll_bar = browser.find_element(By.XPATH, '//div[@id="scroll-vertical"]')
                action = ActionChains(browser)
                action.move_to_element_with_offset(scroll_bar, 0, -scroll_bar.size['height'] / 2).click().perform()
                time.sleep(random.randint(1, 2))

                try:
                    element = browser.find_element(By.XPATH, '//span[@class="onboard-group-name"]')
                    if element.is_displayed():
                        print("Đã tìm thấy element, dừng cuộn.")
                        break
                except:
                    pass

            # Lấy nội dung HTML sau khi cuộn
            page_source = browser.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            messages_data = []
            date_sent_list = []

            # Xử lý HTML để lấy tin nhắn
            blocks = soup.find_all("div", class_="block-date")
            for block in blocks:
                date_element = block.find("span", class_="content")
                if date_element:
                    date_sent = date_element.find("span").get_text(strip=True)
                    date_sent = date_sent.split()[-1]

                messages = block.find_all("div", class_=lambda class_name: class_name and "chat-message wrap-message" in class_name)
                for message in messages:
                    try:
                        sender = message.find("div", class_="truncate").get_text(strip=True).replace("\xa0", " ") if message.find("div", class_="truncate") else pd.NA
                        thoigian = message.find("span-13", class_="card-send-time__sendTime").get_text(strip=True) if message.find("span-13", class_="card-send-time__sendTime") else pd.NA

                        # Lấy nội dung tin nhắn
                        message_content = message.find("div", class_="overflow-hidden")
                        content = ' '.join(message_content.stripped_strings).replace("\n", " ") if message_content else "HÌNH ẢNH"

                        reaction = int(message.find("div", class_="total-reacts").get_text(strip=True)) if message.find("div", class_="total-reacts") else 0

                        messages_data.append({
                            "Người gửi": sender,
                            "Nội dung": content,
                            "Lượt react": reaction,
                            "Thời gian": thoigian
                        })
                        date_sent_list.append(date_sent)

                    except Exception as e:
                        print(f"Error extracting message data: {e}")

            # Tạo DataFrame từ dữ liệu tin nhắn
            df = pd.DataFrame(messages_data)
            df["Ngày gửi"] = date_sent_list

            # Xử lý ngày tháng
            today = datetime.now().strftime("%d/%m/%Y")
            df['Ngày gửi'] = df['Ngày gửi'].replace(to_replace='nay', value=today, regex=True)
            yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
            df['Ngày gửi'] = df['Ngày gửi'].replace(to_replace='qua', value=yesterday, regex=True)
            df['Ngày gửi'] = pd.to_datetime(df['Ngày gửi'], format="%d/%m/%Y")

            # Điền các giá trị thiếu
            df['Người gửi'].fillna(method='ffill', inplace=True)
            df['Thời gian'].fillna(method='bfill', inplace=True)
            df.loc[df['Nội dung'].str.contains("blob:https://chat.zalo.me/", na=False), 'Nội dung'] = "HÌNH ẢNH"

            # Thêm tên group chat vào DataFrame
            df['Room'] = room_name

            # Nối vào DataFrame tổng hợp
            df_full = pd.concat([df_full, df], ignore_index=True)

        return df_full

    except Exception as e:
        print(f"An error occurred while scraping: {str(e)}")

    finally:
        if 'browser' in locals():
            browser.quit()
        print("Browser closed")

# Danh sách các group chat
ds_tengrp = ['CỘNG ĐỒNG ĐẦU TƯ GIÁ TRỊ CÙNG NGA LEE', 'Chứng Khoán Hôm Nay', 'Chứng Khoán Kim Anh TBC', 'SSI-Quỳnh Thư-1529-Tư vấn và thông tin', 'Nhóm tư vấn hỗ trợ anh chị đầu tư chứng khoán', 'TMT.Cộng Đồng Chứng Khoán', 'HaiCK - CLB Chứng Khoán', 'Room (02) TÂM VÕ SSI - ID 1537', 'Trải nghiệm finbox x', 'Vi đạt stocks', 'FABET TƯ VẤN CƠ SỞ VIP', 'Room trải nghiệm 15 TVI Pro Max']
ds_group_id = ['group-item-g315691964407822056', 'group-item-g4869882855979615516', 'group-item-g4361922102574016257', 'group-item-g4499860156686147579', 'group-item-g7028839890132416770', 'group-item-g9053209109070417194', 'group-item-g3067439750607759384', 'group-item-g3807973453928021841', 'group-item-g6874256792402815727', 'group-item-g1943312743074378097', 'group-item-g8947510466865136802', 'group-item-g7662860428534050526']
room_names = ['Nga Lee', 'Chiến Hà', 'Kim Anh', 'Quỳnh Thư', 'Nguyễn Đình Hải', 'Vũ Tâm', 'Quốc Hải Haick', 'Tâm Tâm', 'Phùng Hoa Lý', 'Vi Xuân Đạt', 'Đức mini', 'Thái Hoàng Tvi']

# Chạy hàm và lưu kết quả
df_full = scrape_zalo_profile(ds_tengrp, ds_group_id, room_names)

# Lấy thời gian hiện tại
now = datetime.now()

# Định dạng thời gian thành chuỗi như mong muốn
formatted_time = now.strftime("%Hh%M_%d%m%Y")

# Tạo tên file linh hoạt
file_name = f"Zalo_MessageGrpngoai_{formatted_time}.xlsx"

# Đường dẫn lưu file
output_path = f"C:/Users/Welcome/Documents/VDSC/Task/Task16 - EDA data zalo/Data_RoomNgoai/{file_name}"

# Lưu DataFrame ra file Excel
df_full.to_excel(output_path, index=False)
print(f"Đã lưu file tổng hợp tại {output_path}")
