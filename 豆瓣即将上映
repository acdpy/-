import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import openpyxl

wb = openpyxl.Workbook()
sheet = wb.active

url_coming = 'https://movie.douban.com/cinema/later/qianxinan/'
headers = {
    'user-agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'
}
req = urllib.request.Request(url_coming,headers=headers)
res = urllib.request.urlopen(req)
html = res.read().decode()
soup = BeautifulSoup(html,'html.parser')

sheet.title = soup.find_all('ul',{'class':'tab-hd'})[0].find('li',{'class':'on'}).text

rows = [['电影名字','详细链接','上映时间','类型','地区','热度','预告链接']]
movies = soup.find_all('div',{'class':'tab-bd'})
for movie in movies[0].find_all('div',{'class':'intro'}):
    r =[]
    title = movie.find('a').text
    href = movie.find('a').get('href')
    print(title,href,end=' ')
    r.append(title)
    r.append(href)
    contents = movie.find_all('li')
    for content in contents:
        print(content.text,end=' ')
        r.append(content.text)
    print()
    brief = movie.find('a',{'class':'trailer_icon'})
    try:
        print(brief.text,brief.get('href'))
        r.append(brief.get('href'))
    except AttributeError as e:
        print('暂无预告信息')
    rows.append(r)
for  i in rows:
    sheet.append(i)
print(rows)
wb.save('coming.xlsx')
