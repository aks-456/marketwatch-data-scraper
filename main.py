from bs4 import BeautifulSoup
import xlsxwriter
import requests


tickers = ['adaviation', 'adcb', 'adib', 'bildco', 'taqa', 'adnh', 'adnic', 'adnocdist', 'tkfl', 'adports', 'adxb', 'adnocdrill', 'aglty', 'agthia', 'ajmanbank', 'alfirdous', 'kico', 'alramz', 'salam.bah', 'alsalamsudan', 'asm', 'yahsat', 'aldar', 'alphadhabi', 'amanat', 'amlak', 'apex', 'aram', 'armx', 'emsteel', 'bhmcapital', 'cbi', 'cbd', 'dana', 'dartakaful', 'depa', 'deyaar', 'dewa', 'dfm', 'dic', 'dib', 'aman', 'dnir', 'drc', 'easylease', 'emaardev', 'emaar', 'drive', 'du' 'emiratesnbd', 'erc', 'reit', 'etisalat', 'enbdreit', 'esg', 'eshraq', 
'fertiglobe', 'fh', 'fab', 'fnf', 'fci', 'gfh', 'ghita', 'gchem', 'gulfnav', 'julphar', 'ihc', 'salama', 'ithmr', 'manazel', 'masq', 'methaq', 'multiply', 'rakbank', 'nbq', 'ncc', 'tabreed', 'nmdc', 'watania', 'oic', 'ords', 'outfl', 'palms', 'qholding', 'rakcec', 'rakprop', 'rakwct', 'rapco', 'rpm', 'scidc', 'sib', 'shuaa', 'sudatel', 'takaful.em', 'union', 'upp', 'uab', 'waha']
position = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

text = requests.get('https://www.marketwatch.com/investing/stock/' + tickers[0] + '/download-data').text

soup = BeautifulSoup(text, 'lxml')

data_tags = soup.find_all('td', class_ = ['overflow__cell'])

time_stamp = soup.select('div.cell__content.u-secondary')


workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Ticker')

time_counter = 0 #Counter to make sure time loop stops at 20
line_counter = 1
pos_counter = 20
for time in time_stamp:
    if(time_counter >= 40):
        break
    if(line_counter == 1):
        worksheet.write(position[pos_counter] + str(1), time.text.replace(' ', ''))
        pos_counter -= 1

    if(line_counter == 2):
        line_counter = 0

    time_counter += 1
    line_counter += 1


j = 0
for ticker in tickers:

    text = requests.get('https://www.marketwatch.com/investing/stock/' + tickers[j] + '/download-data').text

    soup = BeautifulSoup(text, 'lxml')

    data_tags = soup.find_all('td', class_ = ['overflow__cell'])


    data_counter = 0 #Counter to ensure data loop stops at 20
    line_counter = 1 #Counter to ensure only every 5th line of data is accessed (for closing price)
    pos_counter = 20 #Positions counter
    for data in data_tags:
        if(data_counter >= 100):
            break
        if(len(data["class"]) != 1):
            continue
            
        if(line_counter == 1):
            data_line = data.text.replace(' ', '')
            data_input = ''
            decimal_tracker = 0 #Checks if number has already been included in data_input to allow only decimal points and not currency symbols
            for character in data_line:
                if(('0' <= character <= '9') or (character == '.' and decimal_tracker == 1)):
                    decimal_tracker = 1 #set decimal_tracker = 1 once a number has been added to data_input
                    data_input += character

            worksheet.write('A' + str(j+2), ticker)
            worksheet.write(position[pos_counter] + str(j+2), data_input)
            pos_counter -= 1

        if(line_counter == 5):
            line_counter = 0

        data_counter += 1
        line_counter += 1
    j += 1

workbook.close()
