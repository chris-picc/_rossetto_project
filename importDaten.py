# input readonly=3D
# f = open('liste.txt', 'r')

# pip install openpyxl
# pip install xlsxWriter

import xlsxwriter
import io

def extract_substring(s, start_str, end_str):
  start_index = s.find(start_str)
  if start_index == -1:
    return ''  # Start string not found
  end_index = s.find(end_str, start_index + len(start_str))
  if end_index == -1:
    return ''  # End string not found after start string
  return s[start_index + len(start_str):end_index]


def verifiz (artikel):
    flag_p = 0
    if artikel == '0132304' or artikel == '0212285' or artikel == '0132307' or artikel == '0136759' or artikel == '0171155' or artikel == '0170579' or artikel == '0145490' or artikel == '0012520' or artikel == '0128900' or artikel == '0012511' or artikel == '0196789' or artikel == '0012556' or artikel == '0012558' or artikel == '0196269' or artikel == '0132301' or artikel == '0147639' or artikel == '0244788' or artikel == '0012509' or artikel == '0012506':
        flag_p = 1
    if artikel == '0254436':
        flag_p = 2
    if artikel == '0219431':
        flag_p = 3        
        
    return flag_p

def preis_neu (preis):
	preis = preis.replace(',','.')
	preis = float(preis)
	# print(type(alltextC[5]))
	preis = preis*6
	preis = round(preis, 3)
	preis = str(preis)
	preis = preis.replace('.',',')
	return preis

def preis_neu_2 (preis):
	preis = preis.replace(',','.')
	preis = float(preis)
	# print(type(alltextC[5]))
	preis = preis/6
	preis = round(preis, 3)
	preis = str(preis)
	preis = preis.replace('.',',')
	return preis
    
def preis_neu_3 (preis):
	preis = preis.replace(',','.')
	preis = float(preis)
	# print(type(alltextC[5]))
	preis = preis/12
	preis = round(preis, 3)
	preis = str(preis)
	preis = preis.replace('.',',')
	return preis
    

f = open('sito_rossetto.mhtml', 'r')

tab = "\t"
absatz = "\n"  
alltext = ""

i = 0
for zeile in f:
  substring = extract_substring(zeile, 'value=3D"', '"')
  
  if "value=3D" in zeile:
      if i==0 or i==1:
        i += 1
      else:
        alltext += substring + tab
        i += 1
      if i==8:
         alltext += absatz
         i=0
       
f.close()

g = open('output.txt', 'w')
g.write(alltext)
g.close()


h = open('output.txt', 'r')

alltextA = []

j = 0
for zeileA in h:
    alltextA.append([zeileA]) # matrice!!!
    j+=1
    
h.close()

file_excel = xlsxwriter.Workbook('Rossetto_Preise_Excel.xlsx')
foglio_excel = file_excel.add_worksheet()

alltextB = alltextA
count=0

for m in range(len(alltextA)):
    # print(m)
    flag_p = 0
    alltextB[count][0] = alltextA[count][0].split('\t')
    alltextC = alltextB[count][0] 
    # ['0179329', 'CAMPIELLO LIGHT GR.350        ', 'Gr.  350', '12,000', '30/04/2025', '1,095', '\n'] 
 # alltextC[0] --> '0179329'
    
    flag_p = verifiz (alltextC[0])
    
    if flag_p == 1:
        alltextC[5] = preis_neu (alltextC[5])
        print(alltextC[5])
        
    if flag_p == 2:
        alltextC[5] = preis_neu_2 (alltextC[5])
        print(alltextC[5])        
        
    if flag_p == 3:
        alltextC[5] = preis_neu_3 (alltextC[5])
        print(alltextC[5])        
        
        
    for n in range(len(alltextB[count][0])):
        if n==0: xy ='A'
        if n==1: xy ='B'
        if n==2: xy ='C'
        if n==3: xy ='D'
        if n==4: xy ='E'
        if n==5: xy ='F'
        if n==6: xy ='G'
        if n==7: xy ='H'
        foglio_excel.write(f'{xy}{m+1}', alltextC[n])
        # print(alltextC[n])
    # print(alltextC[5]) # Preis
    # print(alltextC[5]) # Preis
   
    count+=1
    
file_excel.close()

