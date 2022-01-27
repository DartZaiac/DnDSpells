from bs4 import BeautifulSoup
import urllib.request
import winsound
import docx
# import docx

# import time
from docx import Document
import datetime
from docx.shared import Inches,Pt,Mm,RGBColor
from docx.text.run import Font, Run
from docx.styles.style import _ParagraphStyle
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ROW_HEIGHT
from docx.enum.section import WD_ORIENT
# import docx.shared.Length

# httml='http://www.python.org/'
httml='https://dungeon.su/spells/'
with urllib.request.urlopen(httml) as f:
	listOfSitesOfSpells=[]
	soup = BeautifulSoup (f, 'html.parser')
	for link in soup.find_all('a'):
		if link.get('href').find('/spells/')>=0 and link.get('href').find('http')==-1 and len(link.get('href'))>len('/spells/'):
			if link.get('href').find('1')!=-1 or link.get('href').find('2')!=-1 or link.get('href').find('3')!=-1 or link.get('href').find('4')!=-1 or link.get('href').find('5')!=-1 or link.get('href').find('6')!=-1 or link.get('href').find('7')!=-1 or link.get('href').find('8')!=-1 or link.get('href').find('9')!=-1:
				# print(link.get('href'))
				listOfSitesOfSpells.append('https://dungeon.su'+link.get('href'))
listOfSpels=[[]]
kolOfSpels=len(listOfSitesOfSpells)
# kolOfSpels=10
i=0
# if 1:
try:
	f=open('listOfSpels.txt','r',encoding='utf-8')

	for line in f:
		if line.strip()!='!!!':
			listOfSpels[i].append(line.strip())
		else:
			listOfSpels.append([])
			i=i+1
	listOfSpels.remove([])
	f.close()
	# print(listOfSpels)
	# print(len(listOfSpels))
# except Exception:
	# print(Exception)
except:
	print("No file")
	f=open('listOfSpels.txt','w',encoding='utf-8')
	f.close()
while listOfSpels[len(listOfSpels)-1]=='':
    	listOfSpels.remove[len(listOfSpels)-1]
print(len(listOfSpels))
print(listOfSpels)

if len(listOfSpels)!=kolOfSpels:
	print("Len not same")
	listOfSpels=[[]]
	for i in range(0,kolOfSpels):
	# if 1:
		# i=15
		listOfSpels.append([])
		httml=listOfSitesOfSpells[i]
		with urllib.request.urlopen(httml) as f:
			# print(f.read().decode('utf-8'))
			html_doc= f
			soup = BeautifulSoup (html_doc, 'html.parser')
			if 0:
				print(soup.get_text())
			else:
				pos=soup.title.string.find('(')
				# print(pos)
				txt=soup.title.string[0:pos]
				while txt.rfind(' ')+1==len(txt):
					pos=txt.rfind(' ')
				# print (str(pos)+"   "+str(len(txt)))
					txt=soup.title.string[0:pos]
				title=txt
				
				txt=soup.get_text()
				# print(title+"!")
				pos0=txt.find(title)
				pos0=txt.find(title,pos0+1)
				pos0=txt.find(title,pos0+1)
				pos1=txt.find(']',pos0)+1
				spell_name=txt[pos0:pos1]
				# print("0!"+spell_name+"!")
				listOfSpels[i].append(spell_name)

				pos0=pos1
				pos1=txt.find('Время накладыван',pos0)
				spell_lvl=txt[pos0:pos1]
				# print("1!"+spell_lvl+"!")
				listOfSpels[i].append(spell_lvl)

				pos0=pos1
				pos1=txt.find('Дистанция',pos0)
				spell_time=txt[pos0:pos1]
				# print("2!"+spell_time+"!")
				listOfSpels[i].append(spell_time)
				
				pos0=pos1
				pos1=txt.find('Компоненты',pos0)
				spell_dist=txt[pos0:pos1]
				# print("3!"+spell_dist+"!")
				listOfSpels[i].append(spell_dist)

				pos0=pos1
				pos1=txt.find('Длитель',pos0)
				spell_komponents=txt[pos0:pos1]
				# print("4!"+spell_komponents+"!")
				listOfSpels[i].append(spell_komponents)

				pos0=pos1
				pos1=txt.find('Классы',pos0)
				spell_dlitelnost=txt[pos0:pos1]
				# print("5!"+spell_dlitelnost+"!")
				listOfSpels[i].append(spell_dlitelnost)

				postmp=pos0
				if txt.find('Архетипы',pos0)!=-1:
					pos0=pos1
					pos1=txt.find('Архетипы',pos0)
					# print("!!!"+str(pos1)+"!!!")
					spell_klasses=txt[pos0:pos1]
					# print("6!"+spell_klasses+"!")
					listOfSpels[i].append(spell_klasses)
				else:
					pos0=pos1
					pos1=txt.find('Источник',pos0)
					spell_arhetips=txt[pos0:pos1]
					# print("7!"+spell_arhetips+"!")
					listOfSpels[i].append(spell_arhetips)

				if txt.find('Архети',postmp)!=-1:
					pos0=txt.find('Архетипы',postmp)
					pos1=txt.find('Источник',pos0)
					spell_arhetips=txt[pos0:pos1]
					# print("7!"+spell_arhetips+"!")
					listOfSpels[i].append(spell_arhetips)
				else:
					listOfSpels[i].append("???")
					pass

				pos0=pos1
				pos1=txt.find('»',pos0)+1
				spell_source=txt[pos0:pos1]
				# print("8!"+spell_source+"!")
				listOfSpels[i].append(spell_source)
				
				pos0=pos1
				pos1=txt.find('Поделиться',pos0)-1
				spell_text=txt[pos0:pos1]
				while spell_text.rfind(' ')+1==len(spell_text):
						pos=spell_text.rfind(' ')
						spell_text=spell_text[0:pos]
				# while spell_text.rfind('\n\n')!=-1:
				# 	spell_text.replace('\n\n', '\n')
				spell_text=spell_text[0:len(spell_text)-3]
				# print("9!"+spell_text+"!")
				listOfSpels[i].append(spell_text)
				# print()
				# print(soup.title.name)
				# print(soup.prettify())
	try:
		while 1:
			listOfSpels.remove([])
	except:
		pass
	print(listOfSpels)
	f = open('listOfSpels.txt','w',encoding='utf-8')
	for element in listOfSpels:
		for i in element:
			print(str(i))
			f.write(str(i))
			f.write('\n')
		f.write('!!!\n')
	f.close()
else:
	print("SAME!")
	print()
winsound.Beep(500, 200)
# print(listOfSpels)
listOfClasses=['Бард','Варвар','Воин','Волшебник','Друид','Жрец','Изобретатель','Колдун','Монах','Паладин','Плут','Следопыт','Чародей','Все классы']
fChoice=True
fFlagChoice=0
while fChoice:
	fChoice=not fChoice
	for i in range (0,len(listOfClasses)):
		print(str(i+1)+". "+listOfClasses[i])
	# choiceClass=int(input("Выберите класс:"))
	choiceClass=1
	if choiceClass==1: # 1. Бард

		pass
	elif choiceClass==2:# 2. Варвар
		print("У Варвара нет магии")
		pass
	elif choiceClass==3:# 3. Воин
		while fFlagChoice!=1 and fFlagChoice!=2:
			print("Вы мистический воин?\n1. Да\n2. Нет")
			fFlagChoice= int(input())
		pass
	elif choiceClass==4:# 4. Волшебник

		pass
	elif choiceClass==5:# 5. Друид

		pass
	elif choiceClass==6:# 6. Жрец

		pass
	elif choiceClass==7:# 7. Изобретатель

		pass
	elif choiceClass==8:# 8. Колдун

		pass
	elif choiceClass==9:# 9. Монах

		pass
	elif choiceClass==10:# 10. Паладин

		pass
	elif choiceClass==11:# 11. Плут

		pass
	elif choiceClass==12:# 12. Следопыт

		pass
	elif choiceClass==13:# 13. Чародей
		pass
	elif choiceClass==14:# 13. Все классы
		pass
	else:
		print("Нет такого класса")
		fChoice=True
for i in range (0,len(listOfClasses)):
    		listOfClasses[i]=listOfClasses[i].lower()
listOfLvl=-1
while listOfLvl<0 or listOfLvl>9:
	# listOfLvl=int(input("Уровень заклинания (0-9): "))
	listOfLvl=1
print(listOfClasses[int(choiceClass)-1])
kolClass=0
kolArhetip=0
print("Spisok dlya zapisi")
while listOfSpels[len(listOfSpels)-1]=='':
    	listOfSpels.remove('')
# print(listOfSpels)

doc = Document()
# doc.add_heading('Document Title', 0)
section = doc.sections[-1]
section.orientation = 1
# section.page_width = Inches(10/25.4*29.7)
section.page_width =Mm(297)
# section.page_height = Inches(10/25.4*21)
section.page_height = Mm(210) 
section.top_margin=Mm(5)
section.bottom_margin=Mm(5)
section.left_margin = Mm(5)
section.right_margin =Mm(5)

table0 = doc.add_table(rows=5, cols=4)
hdr_cells = table0.rows[0].cells
table0.rows[0].height=Inches(10/25.4*19.8/2)
table0.rows[1].height=Inches(10/25.4*19.8/2)
table0.rows[2].height_rule = WD_ROW_HEIGHT.EXACTLY
table0.rows[2].height=Inches(10/25.4*0.1)
table0.rows[3].height=Inches(10/25.4*19.8/2)
table0.rows[4].height=Inches(10/25.4*19.8/2)
hdr_cells[0].width=Inches(10/25.4*28.7/4)
hdr_cells[1].width=Inches(10/25.4*28.7/4)
hdr_cells[2].width=Inches(10/25.4*28.7/4)
hdr_cells[3].width=Inches(10/25.4*28.7/4)
t=0
row=0
column=0
fontMain = Font
fontMain.bold=True
fontMain.name='Zaglavniy'
fontMain.size=Pt(8)
fontMain.highlight_color=(0,0,0)
print(fontMain)
for i in listOfSpels:
	if i[1].find(str(listOfLvl))!=-1:
		row=(t//4)
		column=t%4
		hdr_cells = table0.rows[row].cells
		cell = table0.cell(row, column)
		if i[6].find(listOfClasses[3])!=-1 and choiceClass==3 and fFlagChoice==1 and (i[1].find('воплощен')!=-1 or i[1].find('огражден')!=-1):#Мистический рыцарь
			print(i[0])
			cell.font=fontMain
			cell.text=i[0]
			# hdr_cells[column].add_paragraph(i[0], style='List Number')
			# print(i[1])
			# print(i[2])
			# print(i[3])
			# print(i[4])
			# print(i[5])
			# print(i[6])
			if i[6].find(listOfClasses[3])!=-1:
				kolClass+=1
			if i[7].find(listOfClasses[3])!=-1:
				# print(i[7])
				kolArhetip+=1
			# print(i[8])
			try:
				# print(i[9])
				pass
			except:
				pass
			# print()
			t=t+1
			# toDoc(i,doc,t)
		if i[6].find(listOfClasses[int(choiceClass)-1])!=-1 or i[7].find(listOfClasses[int(choiceClass)-1])!=-1:
			# print(i[0])
			
			# cell.text=i[0]
			# cell.paragraphs[0].paragraph_format.space_after=Mm(0)
			# cell.paragraphs[0].paragraph_format.line_spacing=1
			# cell.text='!!!'
			styles = doc.styles
			style=styles.add_style('Main1',WD_STYLE_TYPE.PARAGRAPH)
			paragraph = cell.add_paragraph('','Main1')
			# paragraph.add_run('i[0]').bold = True
			paragraph.add_run('bold').bold = True
			# par_style=_ParagraphStyle
			paragraph = cell.add_paragraph('','Main1')
			paragraph.add_run(i[1]).italic=True			
			# par_style = styles.add_style("Par_Name", WD_STYLE_TYPE.PARAGRAPH)
			# par_font=Font
			# par_font.color=RGBColor(0xB5,0x0C,0xB2)
			# par_font.size=Pt(8)
			# print(par_font.size)
			# styles = doc.styles
			# par_style.font=par_font
			# par_style.name="Main1"
			# 
			# style_font=style.font
			# style_font.size=Pt(8)
			# style_font.color=RGBColor(0xB5,0x0C,0xB2)
			# style.font=style_font
			# style_font.color=
			# print(style.font.size)
			# k=0
			# for stt in styles:
			# 	if stt.name=="Main1":
			# 		print(stt.font.size)
			# 		# stt.font=par_font
			# 		k=k+1
			# print("k="+str(k))
			# par_style.name="Par_Name"
			# cell.paragraphs[0].style='Main1'
			# print(par_font)
			# RGBColor(0xB5,0x0C,0xB2)
			# cell.paragraphs[0].style.font.color=RGBColor(0xB5,0x0C,0xB2)
			# print(cell.paragraphs[0].text)

			# print(cell.paragraphs[0])
			# cell.add_table(2,2)
			# cell.run.bold=True
			# par_for=doc.styles['Normal'].paragraph_format
			# 
			# paragraph.space_before=Pt(1)
			# paragraph.space_after=Pt(1)
			# all_paras=table0.paragraphs
			# par_for=paragraph.paragraph_format.space_after=Pt(0)
			# prior_paragraph = paragraph.insert_paragraph_before(i[0])
			# cell.font.size=Pt(8)
			# cell.font.highlight_color=(0,0,0)
			# hdr_cells[column].add_paragraph(i[0], style='List Number')
			# print(i[1])
			# print(i[2])
			# print(i[3])
			# print(i[4])
			# print(i[5])
			# print(i[6])
			if i[6].find(listOfClasses[int(choiceClass)-1])!=-1:
				kolClass+=1
			if i[7].find(listOfClasses[int(choiceClass)-1])!=-1:
				# print(i[7])
				kolArhetip+=1
			# print(i[8])
			try:
				pass
				# print(i[9])
			except:
				pass
			# print()
			# toDoc(i,doc,t)
			t=t+1
		if t>=1:
			break
		
	
print("Klass: "+str(kolClass)+"  Arhetips: "+str(kolArhetip))
name = str(datetime.datetime.now().time())
name=name.replace(':','_',2)
doc.save('Docs/Example'+name+'.docx')