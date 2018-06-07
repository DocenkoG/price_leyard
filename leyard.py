# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import sys
import configparser
import time
import shutil
import openpyxl                      # Для .xlsx
#import xlrd                          # для .xls
from   price_tools import getCellXlsx, getCell, quoted, dump_cell, currencyType, openX, sheetByName
import csv
import requests, lxml.html



def nameToId(value) :
    result = ''
    for ch in value:
        if (ch != " " and ch != "/" and ch != "\\" and ch != '_' and ch != "," and 
            ch != "'" and ch != "." and ch != "-" and ch != "!" and ch != "@" and 
            ch != "#" and ch != "$" and ch != "%" and ch != "^" and ch != "&" and 
            ch != "*" and ch != "(" and ch != ")" and ch != "[" and ch != "]" and 
            ch != "{" and ch != ":" and ch != '"' and ch != ";"                  ) :
            result = result + ch

    length = len(result)
    if  length > 50 :
        point = int(length/2)
        result = result[:13] + result[point-12:point+13] + result[-12: ]
    
    return result



def getXlsString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]-1
        if item in ('закупка','продажа','цена1') :
            if getCell(row=i, col=j, isDigit='N', sheet=sh).find('Звоните') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCell(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCell(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def getXlsxString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]
        if item in ('закупка','продажа','цена','цена1') :
            if getCellXlsx(row=i, col=j, isDigit='N', sheet=sh).find('Call for Pricing') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCellXlsx(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCellXlsx(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def convert_excel2csv(cfg):
    csvFName  = cfg.get('basic','filename_out')
    priceFName= cfg.get('basic','filename_in')
    sheetName = cfg.get('basic','sheetname')
    
    log.debug('Reading file ' + priceFName )
    sheet = sheetByName(fileName = priceFName, sheetName = sheetName)
    if not sheet :
        log.error("Нет листа "+sheetName+" в файле "+ priceFName)
        return False
    log.debug("Sheet   "+sheetName)
    out_cols = cfg.options("cols_out")
    in_cols  = cfg.options("cols_in")
    out_template = {}
    for vName in out_cols :
         out_template[vName] = cfg.get("cols_out", vName)
    in_cols_j = {}
    for vName in in_cols :
         in_cols_j[vName] = cfg.getint("cols_in",  vName)

    csvFNameRUR =csvFName[:-4]+'_RUR'+csvFName[-4:]
    csvFNameEUR =csvFName[:-4]+'_EUR'+csvFName[-4:]
    csvFNameUSD =csvFName[:-4]+'_USD'+csvFName[-4:]
    outFileRUR = open( csvFNameRUR, 'w', newline='', encoding='CP1251', errors='replace')
    outFileEUR = open( csvFNameEUR, 'w', newline='', encoding='CP1251', errors='replace')
    outFileUSD = open( csvFNameUSD, 'w', newline='', encoding='CP1251', errors='replace')
    csvWriterRUR = csv.DictWriter(outFileRUR, fieldnames=out_cols )
    csvWriterEUR = csv.DictWriter(outFileEUR, fieldnames=out_cols )
    csvWriterUSD = csv.DictWriter(outFileUSD, fieldnames=out_cols )
    csvWriterRUR.writeheader()
    csvWriterEUR.writeheader()
    csvWriterUSD.writeheader()

    recOut  ={}
    for i in range(2, sheet.max_row +1) :                                # xlsx
#   for i in range(2, sheet.nrows) :                                     # xls
        i_last = i
        try:
            impValues = getXlsxString(sheet, i, in_cols_j)               # xlsx
            #impValues = getXlsString(sheet, i, in_cols_j)               # xls
            #print( impValues )
            #if sheet.cell( row=i, column=in_cols_j['код_']).fill.fgColor.type=='indexed':   # подгруппа
                #print(sheet.cell( row=i, column=in_cols_j['код_']).fill.fgColor.indexed)
                #col2 = sheet.cell( row=i, column=in_cols_j['код_']).value.strip()
                #t = col2.rpartition(' ')
                #brand  = t[2]
                #subgrp = t[0]
                #continue
            if impValues['цена1']=='0': # (ccc.value == None) or (ccc2.value == None) :   # Пустая строка
                #print( 'Пустая строка. i=',i, impValues )
                continue
            else :                                                         # Обычная строка
                for outColName in out_template.keys() :
                    shablon = out_template[outColName]
                    for key in impValues.keys():
                        if shablon.find(key) >= 0 :
                            shablon = shablon.replace(key, impValues[key])
                    if (outColName == 'закупка') and ('*' in shablon) :
                        p = shablon.find("*")
                        vvv1 = float(shablon[:p])
                        vvv2 = float(shablon[p+1:])
                        shablon = str(round(vvv1 * vvv2, 2))
                    elif (outColName=='код') :
                        shablon = nameToId(shablon)
                    recOut[outColName] = shablon.strip()

            if   'RUR'==recOut['валюта'] :
                       csvWriterRUR.writerow(recOut)
            elif 'USD'==recOut['валюта'] :
                       csvWriterUSD.writerow(recOut)    
            elif 'EUR'==recOut['валюта'] :
                       csvWriterEUR.writerow(recOut)    
            
        except Exception as e:
            print(e)
            if str(e) == "'NoneType' object has no attribute 'rgb'":
                pass
            else:
                log.debug('Exception: <' + str(e) + '> при обработке строки ' + str(i) +'.' )

    log.info('Обработано ' +str(i_last)+ ' строк.')
    outFileRUR.close()
    outFileUSD.close()



def config_read( cfgFName ):
    cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
    if  os.path.exists('private.cfg'):     
        cfg.read('private.cfg', encoding='utf-8')
    if  os.path.exists(cfgFName):     
        cfg.read( cfgFName, encoding='utf-8')
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    return cfg



def is_file_fresh(fileName, qty_days):
    qty_seconds = qty_days *24*60*60 
    if os.path.exists( fileName):
        price_datetime = os.path.getmtime(fileName)
    else:
        log.error('Не найден файл  '+ fileName)
        return False

    if price_datetime+qty_seconds < time.time() :
        file_age = round((time.time()-price_datetime)/24/60/60)
        log.error('Файл "'+fileName+'" устарел!  Допустимый период '+ str(qty_days)+' дней, а ему ' + str(file_age) )
        return False
    else:
        return True



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def processing(cfgFName):
    log.info('----------------------- Processing '+cfgFName )
    cfg = config_read(cfgFName)
    filename_out = cfg.get('basic','filename_out')
    filename_in  = cfg.get('basic','filename_in')
    
    convert_excel2csv(cfg)
    if os.name == 'nt' :
        folderName = os.path.basename(os.getcwd())
        foutRUR = filename_out[:-4]+'_RUR'+filename_out[-4:]
        foutEUR = filename_out[:-4]+'_EUR'+filename_out[-4:]
        foutUSD = filename_out[:-4]+'_USD'+filename_out[-4:]

        if os.path.exists(foutRUR)  : shutil.copy2(foutRUR , 'c://AV_PROM/prices/' + folderName +'/'+foutRUR)
        if os.path.exists(foutEUR)  : shutil.copy2(foutEUR , 'c://AV_PROM/prices/' + folderName +'/'+foutEUR)
        if os.path.exists(foutUSD)  : shutil.copy2(foutUSD , 'c://AV_PROM/prices/' + folderName +'/'+foutUSD)
        if os.path.exists('python.log')  : shutil.copy2('python.log',  'c://AV_PROM/prices/' + folderName +'/python.log')
        if os.path.exists('python.log.1'): shutil.copy2('python.log.1','c://AV_PROM/prices/' + folderName +'/python.log.1')
    


def main( dealerName):
    """ Обработка прайсов выполняется согласно файлов конфигурации.
    Для этого в текущей папке должны быть файлы конфигурации, описывающие
    свойства файла и правила обработки. По одному конфигу на каждый 
    прайс или раздел прайса со своими правилами обработки
    """
    make_loger()
    log.info('          '+dealerName )

    if  os.path.exists('getting.cfg'):     
        cfg = config_read('getting.cfg')
        filename_new = cfg.get('basic','filename_new')
        
        rc_download = False
        if cfg.has_section('download'):
            rc_download = download(cfg)
        if rc_download==True or is_file_fresh( filename_new, int(cfg.get('basic','срок годности'))):
            pass
        else:
            return
        
    for cfgFName in os.listdir("."):
        if cfgFName.startswith("cfg") and cfgFName.endswith(".cfg"):
            processing(cfgFName)


if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( myName)
