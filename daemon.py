from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options
import pandas as pd
import time

options = webdriver.FirefoxOptions()

df = pd.read_csv("parameters.csv")
account = df["value"][0]
pwd = df["value"][1]
conlentm = df["value"][5]
loopcir = df["value"][9]
path_gecko = df["value"][10]
#hdless = df["value"][10]

options = Options()
options.add_argument('zh_CN')
options.add_argument('User-Agent:Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36')
#options.headless = hdless
options.add_argument('--headless')
options.add_argument('--disable-gpu')
browser = webdriver.Firefox(executable_path='{}'.format(path_gecko))
browser.get("http://wljx.gsau.edu.cn/meol/index.do")
browser.find_element_by_xpath('//*[@id="loginbtn"]').click()
time.sleep(3)
browser.find_element_by_xpath('//*[@id="userName"]').send_keys(account)
time.sleep(1)
browser.find_element_by_xpath('//*[@id="passWord"]').send_keys(pwd)
time.sleep(1)
browser.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[3]/div/form/div[2]/input').click()
print("登录成功！")
time.sleep(4)
#切换tab
browser.find_element_by_xpath("/html/body/div[1]/div[3]/div/div/div/div[3]/div[3]/div/div/div[2]/ul/li[14]/div[2]/p[1]/a").click()
time.sleep(1)
browser.switch_to.window(browser.window_handles[-1])
browser.find_element_by_xpath('/html/body/div[3]/div[1]/div[2]/div[2]/ul/li[3]/a').click()
time.sleep(1)
#进入学习界面
browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div[2]/div[2]/ul[3]/li[1]/a/span').click()
time.sleep(1)
print("打开学习界面成功！")

#视频执行模块
def video_opt (j):
    video_webpage = 'window.open("{}")'.format(df["value"][j+2])
    browser.execute_script(video_webpage)
    #切换到视频播放界面 打开第一个视频
    browser.switch_to.window(browser.window_handles[-1])
    time.sleep(1)
    ActionChains(browser).move_by_offset(500,500).click().perform()
    time.sleep(int(df["value"][j+6]))
    print("第{}个视频学习完毕".format(j+1))
    return

#循环开始处
for i in range(0,int(loopcir)):
    print("执行第{}个循环".format(i+1))
    browser.switch_to.window(browser.window_handles[-1])
    #打开电子讲稿学习
    browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div[2]/div[2]/ul[3]/li[1]/ul/li[1]/a/span').click()
    time.sleep(int(conlentm))
    #打开视频资源学习
    browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div[2]/div[2]/ul[3]/li[1]/ul/li[2]/a/span').click()
    time.sleep(1)
    #打开视频
    for j in range(0,3):
        video_opt(j)
        time.sleep(1)
    #关闭视频tabs
    browser.close()
    browser.switch_to.window(browser.window_handles[-1])
    browser.close()
    browser.switch_to.window(browser.window_handles[-1])
    browser.close()
    print("循环结束，开始下一轮")

print("全部循环已结束！")
