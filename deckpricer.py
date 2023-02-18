import os
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font


def CookSoup(url):
    agent = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36", 'referer':'https://www.google.com/'}
    data = requests.get(url, headers=agent)
    return BeautifulSoup(data.text, 'html.parser')

def order(unordered_lst):
    ordered_cardlst = []
    while (unordered_lst != []):
        max_val = -1
        maxind = -1
        for i in range (len(unordered_lst)):
            if max_val < unordered_lst[i][1]:
                max_val = unordered_lst[i][1]
                maxind = i
        
        max_val = round(max_val, 2)
        
        ordered_cardlst.append((unordered_lst[maxind][0], max_val))
        
        unordered_lst.pop(maxind)

    return ordered_cardlst

def create_xlsx(cardlst, Filename):    
    Column = 'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z'.split(' ')
    
    workbook = Workbook()
    sheet = workbook.active
    sheet.column_dimensions['A'].width = 25
    sheet.column_dimensions['B'].width = 14
    
    sheet['A1'] = 'Total'
    sheet ['A1'].font = Font(bold=True)
    sheet['B1'] = '=SUM(B2:B110)'
    sheet ['B1'].font = Font(bold=True)

    
    for i in range (len(cardlst)):
        for ii in range (len(cardlst[i])):
            sheet [Column[ii] + str(i+2)] = cardlst[i][ii]
            sheet [Column[ii] + str(i+2)].data_type = 'n'
    
    workbook.save(filename=f'./deckfiles/{Filename}_prices.xlsx')


# ================================================================== MAIN ==================================================================
for Filename in os.listdir('./deckfiles'):
    if '.cod' not in Filename: continue
    
    Filename = Filename.replace('.cod', '')
    if f'{Filename}_prices.xlsx' in os.listdir('./deckfiles'):
        continue
    
    with open(f'./deckfiles/{Filename}.cod', 'r') as f:
        cards = f.readlines()

    total = 0
    cardlst = []
    for card in cards:
        if '<card number=' in card:
            card = card.split('name="')[1].split('"/>')[0]
            
            if card.lower() in ['mountain', 'island', 'forest', 'swamp', 'plains']:
                continue
            
            payload = card.replace(' ', '+').replace(',', '%2C').replace("'", '%27')
            if '//' in payload:
                payload = payload.split('//')[0].strip()
            
            url = f'https://www.cardkingdom.com/catalog/search?search=header&filter%5Bname%5D={payload}'
            
            var = str(CookSoup(url))
            var = var.split('class="itemContentWrapper"')
            
            min_val = 999.99

            for i in range (len(var)):
                if 'class="productDetailTitle">' in var[i]:
                    ph = var[i].split('class="productDetailTitle">')[1].split('</span>')[0]
                    if '</a>' in ph: ph = ph.split('">')[1].split('</a>')[0]
                    name = ph
                    
                    ph = var[i].split('class="stylePrice">')[1].split('</span>')[0].strip()
                    ph = ph.replace('$', '').replace(',', '').strip()
                    price = float(ph)      

                    if ('art card' not in name.lower()) and ('not tournament legal' not in name.lower()):
                        if price < min_val:
                            min_val = price
                        
                    
            if min_val == 999.99:
                min_val = 0
                print (f"{card}: N/A")
                total += min_val
                cardlst.append((card, min_val))
            else:
                min_val = round(min_val, 2)
                print (f"{card}: ${min_val}")
                total += min_val
                cardlst.append((card, min_val))
        
    total = round(total, 2)

    print (f'Total: ${total}')

    create_xlsx(order(cardlst), Filename)



