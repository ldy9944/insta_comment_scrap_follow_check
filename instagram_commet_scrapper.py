from selenium import webdriver
import time
from openpyxl import Workbook
import pandas as pd
import instaloader

# Excel シート作成
wb = Workbook(write_only=True)
ws = wb.create_sheet()

# URL, アカウント情報指定
url = input('instagramの投稿URLを入力してください。\n')
insta_id = input('instagramのIDを入力してください。\n')
insta_pw = input('instagramのパスワードを入力してください。\n')

# ブラウザーを開き、最大サイズに変更
driver = webdriver.Chrome()
driver.maximize_window()
driver.get(url)
time.sleep(1)

# 該当の投稿にすぐアクセスする場合はボタンを押してログイン画面へ
# ログイン画面が先に表示されるとすぐログイン
while True:
    try:
        login_btn = driver.find_element_by_class_name('ENC4C')
    except:
        break

    if login_btn is not None:
        try:
            login_btn.click()
            time.sleep(1)
        except:
            break

# アカウント情報を入力の上、ログイン
# アカウント情報の保存はしない
driver.find_element_by_css_selector('#loginForm > div > div:nth-child(1) > div > label > input').send_keys(insta_id)
driver.find_element_by_css_selector('#loginForm > div > div:nth-child(2) > div > label > input').send_keys(insta_pw)
time.sleep(1)
driver.find_element_by_css_selector('#loginForm > div > div:nth-child(3) > button').click()
time.sleep(3)
driver.find_element_by_class_name('cmbtv').click()

time.sleep(3)

# コメント追加表示ボタンを押し、全てのコメントを表示
while True:
    try:
        button = driver.find_element_by_css_selector('div.EtaWk > ul > li > div > button')
        pass
    except:
        pass

    if button is not None:
        try:
            button.click()
            time.sleep(1)
        except:
            break

time.sleep(1)

# ユーザー名とコメントをリスト化

container = driver.find_elements_by_class_name('Mr508 ')

id_f = list()
rp_f = list()

for c in container:
    id = c.find_element_by_class_name('_6lAjh')

    id_f.append(id.text)

    rp = c.find_element_by_class_name('MOdxS ')
    rp_f.append(rp.text)

# フォロー確認
all_follow = list()
now_follow = list()

L = instaloader.Instaloader()

L.login(insta_id, insta_pw)

profile = instaloader.Profile.from_username(L.context, insta_id)

count = 0
for followee in profile.get_followers():
    all_follow.append(followee.username)
    file = open("prada_followers.txt", "a+")
    file.write(all_follow[count])
    file.write("\n")
    file.close()
    print(all_follow[count])
    count = count + 1

for i in id_f:
    if i in all_follow:
        now_follow.append('O')
    else:
        now_follow.append('X')

# Excel作成、プログラムを終了
data = {"ユーザー名": id_f, 'フォロー': now_follow, "コメント": rp_f}


df = pd.DataFrame(data)
df.to_excel('result.xlsx')

driver.quit()
