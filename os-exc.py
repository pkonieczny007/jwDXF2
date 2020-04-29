import os 
import fdxf
import openpyxl

n=0


wb = openpyxl.load_workbook('1.xlsx')
sheet = wb.active
sheet.title = 'wykaz'





for filename in os.listdir():
    if filename.endswith('.dxf'):
        try:
            n+=1
            wys=fdxf.wys(filename)
            szer=fdxf.sze(filename)
            print(n, filename, wys, szer)

            a="A"+str(n)
            b="B"+str(n)
            c="C"+str(n)
            d="D"+str(n)
            
            sheet[a]=str(n)
            sheet[b]=str(filename)
            sheet[c]=str(wys)
            sheet[d]=str(szer)
        except:            
            a="A"+str(n)
            b="B"+str(n)
            c="C"+str(n)
            d="D"+str(n)
            e="E"+str(n)
            sheet[a]=str(n)
            sheet[b]=str(filename)
            sheet[c]=str(wys)
            sheet[d]=str(szer)
            sheet[e]='blad'

        wb.save('1.xlsx')



wb.save('1.xlsx')
wb.close()
