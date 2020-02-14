# Soal 2
import xlsxwriter
book = xlsxwriter.Workbook('SegitigaExcel.xlsx')
sheet = book.add_worksheet('Sheet 1')


def segitigaExcel(kata):
    syarat = [1]
    awal = 1
    kata = kata.replace(' ', '')
    for i in range(2, len(kata)):
        awal += i
        syarat.append(awal)
    if len(kata) in syarat:
        stringPenampung = ''
        for i in range(len(kata)):
            kata = kata.replace(' ', '')
            x = i + 1
            for j in range(i):
                i += j
            stringPenampung += kata[i:i+x]
            stringPenampung += '\n'
            if i > len(kata):
                break
        dataList = []
        listString = stringPenampung.split('\n')
        for i in listString:
            if len(i) > 0:
                dataList += [i]
        row = 0
        for char in dataList:
            for i in range(len(char)):
                sheet.write(row, i, char[i])
            row += 1
        book.close()
    else:
        print('Mohon maaf, jumlah karakter tidak memenuhi syarat membentuk pola.')


segitigaExcel('Purwadhika')
segitigaExcel('Purwadhika Startup and Coding School @BSD')
segitigaExcel('kode')
segitigaExcel('kode python')
segitigaExcel('Lintang')
