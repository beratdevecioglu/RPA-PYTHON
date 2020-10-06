import rpa as r, xlsxwriter

r.init()


workbook = xlsxwriter.Workbook('İhale.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('B:B', 100)
worksheet.set_column('C:C', 75)
worksheet.set_column('D:D', 25)
worksheet.write('B1', 'İhale Adı')
worksheet.write('C1', 'İdare Adı')
worksheet.write('D1','İhale Onaylanma Tarihi')



r.url('https://ekap.kik.gov.tr/EKAP/Ortak/IhaleArama/index.html')
r.wait(8)

r.click('/html/body/div/div[1]/div/div/div[4]/div[2]/input')
r.type('/html/body/div/div[1]/div/div/div[4]/div[2]/input','bilgisayar')
r.wait(1)
r.click('/html/body/div/div[1]/div/div/div[4]/div[4]/div[3]/select')
r.wait(1)
r.type('/html/body/div/div[1]/div/div/div[4]/div[4]/div[3]/select','i')
r.wait(1)
r.click('pnlFiltreBtn')
r.wait(3)
total_items = r.count('div.col-sm-12')
print("Toplam İhale Sayısı =",total_items)

r.wait(3)






for x in range(0, total_items):
    
    print(x+1,". İhale") 
    worksheet.write(x+1, 0, x+1)    
    ihale_adi = r.read('//*[@id="sonuclar"]/div[' + str(x+1) + ']/div/div/div/div[2]/div/div/p[1]')
    print(ihale_adi) 
    worksheet.write(x+1, 1, ihale_adi)
    idare_adi = r.read('//*[@id="sonuclar"]/div[' + str(x+1) + ']/div/div/div/div[3]/div/div/p')
    print(idare_adi)
    worksheet.write(x+1, 2, idare_adi)
    ihale_detayi = r.read('//*[@id="sonuclar"]/div[' + str(x+1) + ']/div/div/div/div[2]/div/div/p[2]')
    print(ihale_detayi)
    worksheet.write(x+1, 3, ihale_detayi)
    
    

workbook.close()
r.close()