# -*- coding: utf-8 -*-
import pandas
import math
import datetime
import sys,os
import shapefile

start = datetime.datetime.now()


path = u'C:\\Users\\HPZ640\\Desktop\\WORKLOAD\\AF\\BASKENT\\PROJE\\ENERJISADAN_GELENLER\\PROJELER - Copy'

##TODO Check for the most updated inventory lookup values
lookup_value = [u'INDIRICI&DAGITICI_MERKEZ SINIRI'
u'INDIRICI&DAGITICI_MERKEZLER'
u'TRAFO_BINA_SINIRI'
u'TRAFO_BINASI'
u'DUT'
u'TOPRAKLAMA_DEGERLERI'
u'YGYG_TRAFO'
u'YGAG_TRAFO'
u'HUCRE'
u'GERILIM_TRAFO_GRUBU'
u'AKIM_TRAFO_GRUBU'
u'DAHILI_YUK_AYIRICI'
u'HARICI_YUK_AYIRICI'
u'KESICI'
u'TEKRAR_KAPAMALI_KESICI'
u'AYIRICI'
u'YGSIGORTA_GRUBU'
u'DIREK'
u'IZOLATOR'
u'SOKAK_AYDINLATMA'
u'BARA'
u'YG_HAT'
u'YG_KABLO'
u'AG_HAT'
u'AG_KABLO'
u'REKORTMAN_HATTI'
u'REDRESOR'
u'AKU_GRUBU'
u'PARAFUDR GRUBU'
u'ENERJI_ANALIZORU'
u'ROLE'
u'AGPANO'
u'SDK'
u'AGSIGORTA_GRUBU'
u'SALTER'
u'YG-AG ŞÖNT KONDANSATÖR BANKI'
u'ŞÖNT REAKTÖR'
u'JENERATOR'
u'AGD'
u'KABLO MUF VE BASLIGI'
u'DIREK PARAFUDR GRUBU'
u'DIREK AYIRICI'
u'DIREK_YGSIGORTA_GRUBU'
u'DIREK_YG_AG_SONT_KONDANSATOR_BANKI'
u'DIREK REAKTOR'
u'KONTROL AMAÇLI VE AYDINLATMA SAYAÇLARI']


def return_checked_values(var):
    if isinstance(var, str):
        print var
        return False
    if isinstance(var, float):
        if math.isnan(var):
            return False
        else:
            return True
    if isinstance(var, unicode):
        if var == u'KOORDINAT_X' or var== u'KOORDINAT_Y' :
            return False
        else:
            return True

X = []
Y = []
Error = []
pyp_check = []
res = []
for r,d,f in os.walk(path):
    for en, file in enumerate(f):
        if (file.find(".xls") != -1  or file.find(".XLSX") != -1 or file.find(".xlsx") != -1 or file.find(".XLS") != -1 ) and not file.__contains__("~"):
            try:
                print os.path.join(r,file)
                fname = os.path.join(r, file)
                folder_pyp_name = r.split("\\")[10]
                print folder_pyp_name
                print fname
                a = pandas.ExcelFile(fname)
                temp = []
                for i in a.sheet_names:

                    b = a.parse(i)

                    koordinat_X_location = [[row, en] for en, row in enumerate(b.iloc[:, 0]) if row == u'KOORDINAT_X']
                    TEKHAT = [[row, en] for en, row in enumerate(b.iloc[:, 0]) if row == u'TEKHAT']

                    if len(koordinat_X_location) > 0 and len(TEKHAT) == 0:
                        koordinat_Y_location = [[row, en] for en, row in enumerate(b.iloc[:, 0]) if row == u'KOORDINAT_Y']
                        pano_numarasi_location = [[row, en] for en, row in enumerate(b.iloc[:, 0]) if row == u'PANO_NUMARASI']
                        hucre_numarasi_location = [[row, en] for en, row in enumerate(b.iloc[:, 0]) if row == u'HUCRE_NO']
                        name_location = [b.iloc[row[1] - 1, 0] for row in koordinat_X_location]


                        for en, row in enumerate(koordinat_X_location):
                            xx = b.iloc[row[1], :]
                            yy = b.iloc[koordinat_Y_location[en][1], :]
                            xxx = [yaw for yaw in xx if return_checked_values(yaw)]
                            yyy = [yaw for yaw in yy if return_checked_values(yaw)]
                            temp.append([name_location[en],xxx,yyy])
                            #print [name_location[en],xxx,yyy]
                        pano = [yaw for yaw in b.iloc[pano_numarasi_location[0][1], :] if return_checked_values(yaw)]
                        hucre = [yaw for yaw in b.iloc[hucre_numarasi_location[0][1], :] if return_checked_values(yaw)]
                        temp.append(pano)
                        temp.append(hucre)
                        temp2 = []
                        for t in temp:
                            if len(t)>1:
                                if t[0] in [u'PANO_NUMARASI', u'HUCRE_NO']:
                                    temp2.extend([t[0], len(t) - 1])
                                else:
                                    temp2.extend([t[0], len(t[1])])
                            else:
                                temp2.extend([t[0], 0])
                        res.append([folder_pyp_name,temp2])
                        proje_numarasi_location = [[row, en] for en, row in enumerate(b.iloc[:, 0]) if row == u'PROJE_NUMARASI']
                        proje_numarasi = [b.iloc[row[1], 1] for row in proje_numarasi_location]
                        # pyp_check.append([folder_pyp_name, "::".join(list(filter(lambda x: not math.isnan(x), proje_numarasi)))])
                        print folder_pyp_name
                        X_temp = [float(bbb.replace(',','.')) for bbb in list(set([str(j) for i in [row[1:] for row in b._values if row[0] == u'KOORDINAT_X'] for j in i])) if bbb != 'nan']
                        Y_temp = [float(bbb.replace(',','.')) for bbb in list(set([str(j) for i in [row[1:] for row in b._values if row[0] == u'KOORDINAT_Y'] for j in i])) if bbb != 'nan']
                        if len(X_temp) == len(Y_temp):
                            X.append(X_temp)
                            Y.append(Y_temp)
                        else:
                            Error.append(fname)
            except BaseException as be:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                Error.append(fname)
                print(exc_type,"Exception Line Number :::::::>" ,exc_tb.tb_lineno)
                print be.message

pass

print pyp_check
end = datetime.datetime.now()

print "The script has been consumed in : ", end-start
