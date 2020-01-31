import bs4
import xlwings as xw
from selenium import webdriver

browser = webdriver.Chrome('../webdriver/chromedriver')
browser.get('https://www.billboard.com/charts/hot-100')  # selenium open web page

top100s = bs4.BeautifulSoup(browser.page_source, "html.parser")

weekSelector = top100s.select(".date-selector__button")
# print(weekSelector[0].getText())
thisWeek = weekSelector[0].getText()  # get last date of this week (always Saturday)

# create excel file
app = xw.App(visible=True, add_book=False)
workbook = app.books.add()
sheet = workbook.sheets[0]
sheetTitle = [['Rank', 'Name', 'Artist', 'Last Week', 'Peak', 'Duration in Chart']]
sheet.range('A1').value = sheetTitle
sheet.range('B1').column_width = 14  # set column width of song names
sheet.range('C1').column_width = 12  # set column width of artists
###

songItems = top100s.select(".chart-list__element")  # get list items of top100 songs
songNames = top100s.select(".chart-list__elements .chart-element__information__song")
artists = top100s.select(".chart-list__elements .chart-element__information__artist")
lastWeeks = top100s.select(".chart-element__metas .text--last")
peaks = top100s.select(".chart-element__metas .text--peak")
durations = top100s.select(".chart-element__metas .text--week")

# print(len(songNames))
# print(len(artists))
# print(len(lastWeeks))
# print(len(peaks))
# print(len(durations))

ranks = 1
for songs in songItems:
    sheet[ranks, 0].value = ranks  # set rank of songs in excel
    sheet[ranks, 1].value = songNames[ranks - 1].getText()  # set names of songs
    sheet[ranks, 2].value = artists[ranks - 1].getText()  # set artists of songs
    sheet[ranks, 3].value = lastWeeks[ranks - 1].getText()  # set last week ranks of songs
    sheet[ranks, 4].value = peaks[ranks - 1].getText()  # set peak ranks of songs
    sheet[ranks, 5].value = durations[ranks - 1].getText()  # set durations in chart of songs

    ranks += 1

workbook.save(thisWeek + ".xlsx")  # save excel file name based on weeks
