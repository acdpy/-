import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import openpyxl

wb = openpyxl.Workbook()
sheet = wb.active

url_now = r'https://movie.douban.com/cinema/nowplaying/qianxinan/'
headers = {
    'user-agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'
}
req = urllib.request.Request(url_now,headers=headers)
res = urllib.request.urlopen(req)
html = res.read().decode()
soup = BeautifulSoup(html,'html.parser')

sheet.title = soup.find('h2').text
rows = [['电影名字','影评','导演','时长','上映地区','详细信息']]
movies = soup.find('div',{'id':'nowplaying'})
for movie in movies.find_all('li'):
    r = []
    title = movie.get('data-title')
    score = movie.get('data-score')
    actors = movie.get('data-actors')
    duration = movie.get('data-duration')
    region = movie.get('data-region')
    if score:
        print(title,region,score,duration,actors,end=' ')
        for i in (title,region,score,duration,actors):
            r.append(i)
        for detail in movie.find_all('a',{'data-psource':'poster'}):
            print(detail.get('href'))
            r.append(detail.get('href'))
    rows.append(r)
for i in rows:
    if not i:
        continue
    sheet.append(i)
print(rows)
wb.save('playing.xlsx')



