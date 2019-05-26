#! /usr/bin/env/python
# -#*- coding: utf-8 -*-
import os, sys, getopt, codecs, csv, MySQLdb, platform
from mmap import mmap,ACCESS_READ
from xlrd import open_workbook, xldate_as_tuple

# Define constants

NALET_OUT = ''
PUTNICI_OUT = ''
DB_HOST = 'localhost'
DB_USER = 'fikus'
DB_PASS = 'qk83xasn'
DB_DATABASE = 'eth'
START_DIR = u'/media/shared_disk/mega tablica'
ERROR_FILE = START_DIR + 'mega_error.log'

# Functions    
def isNumber(s):
# Check if a string could be a number
    try:
        float(s)
        return True
    except ValueError:
        return False

def getMonth(f):
# Izvuci mjesec iz imena datoteke u formatu "1_sijecanj.xls"

    temp = os.path.basename(f)
    temp = temp.split('_')
    mjesec = int(temp[0])
    return mjesec

def getYear(f):
# Izvuci godinu iz path
    f = f.split('/')
    godina = f[-2]
    return godina

def databaseVersion(cur):
# Print Mysql database version
    try:
        cur.execute("SELECT VERSION()")
        result = cur.fetchone()
    except MySQLdb.Error, e:
        try:
            print "MySQL Error [%d]: %s]" % (e.args[0], e.args[1])
        except IndexError:
            print "MySQL  Error: %s" % (e.args[0], e.args[1])

    print "MySQL database version: %s" % result

def getQuery(cur, sql_query):
# Perform passed query on passed database
    try:
        cur.execute(sql_query)
        result = cur.fetchall()
    except MySQLdb.Error, e:
        try:
            print "MySQL Error [%d]: %s]" % (e.args[0], e.args[1])
        except IndexError:
            print "MySQL  Error: %s" % (e.args[0], e.args[1])
    return result

def getFiles():
    
    files = []

    # Find subdirectories
    for i in [x[0] for x in os.walk(START_DIR)]:

        if (i != '.' and isNumber(os.path.basename(i))):

            # Find files in subdirectories
            for j in [y[2] for y in os.walk(i)]:

                # For every file in file list
                for y in j:
                    fn, fe = os.path.splitext(y)
                    is_mj = fn.split("_")
                    if(fe == '.xls' and y.find('_') and isNumber(is_mj[0])):
                        mj = fn.split('_')
                        # if START_DIR="./", i.lstrip('./'), else i
                        print "%s" % (i + "/" + y)
                        files.append(i + "/" + y)
                    
    # Sort list chronologically
    files.sort(key=lambda x: getMonth(x))
    files.sort(key=lambda x: getYear(x))

    print "%s" % files[0]
    return files

def errhandle(f, datum, var, vrijednost, ispravka = "NULL"):
# Get error information, print it on screen and write to error.log

    # f = unicode(str(f), 'utf-8')
    datum = unicode(str(datum), 'utf-8')
    var = unicode(str(var), 'utf-8')
    try:
        vrijednost = unicode(str(vrijednost.decode('utf-8')), 'utf-8')
    except UnicodeEncodeError:
        vrijednost = vrijednost
    ispravka = unicode(str(ispravka), 'utf-8')
    
    err_f = codecs.open(ERROR_FILE, 'a+', 'utf-8')
    line = f + ": " + datum + " " + var + "='" + vrijednost\
                    + "' Ispravka='" + ispravka + "'"

    #print "%s" % line

    err_f.write(line)
    err_f.close()
    
def readxlsfile(files, sheet, piloti, tehnicari, helikopteri):
# Read xls file and return a list of rows

    data = []
    nalet = []
    putn = []
    id_index = 0
    
    # For every file in list
    for f in files:
        print "%s" % f
        temp = f.split('/')
        godina = str(temp[-2])
        temp = os.path.basename(f).split('_')
        mjesec = str(temp[0])

        workbook = open_workbook(f.encode(sys.getfilesystemencoding()))
        sheet = workbook.sheet_by_name('UPIS')
	
        # For every row that doesn't contain '' or 'POSADA' or 'dan' etc...
        for ri in range(sheet.nrows):
            if sheet.cell(ri,1).value!=''\
               and sheet.cell(ri,2).value!='POSADA'\
               and sheet.cell(ri,1).value!='dan'\
               and (sheet.cell(ri,2).value!=''):
                
                temp = sheet.cell(ri, 1).value
                temp = temp.split('.')
                dan = temp[0]
                
                # Datum
                datum = "'" + godina + "-" + mjesec + "-" + dan + "'"
                
                # Kapetan
                kapetan = ''
                kapi=''
                if sheet.cell(ri, 2).value == "":
                    kapetan = "NULL"
                else:
                    kapetan = sheet.cell(ri, 2).value
                    if kapetan[-1:] == " ":
                        errhandle(f, datum, 'kapetan', kapetan, kapetan[-1:])
                        kapetan = kapetan[:-1]
                    if(kapetan):
                        try:
                            kapi = [x[0] for x in piloti if x[2].lower() == kapetan]
                            kapi = kapi[0]
                        except ValueError:
                            errhandle(f, datum, 'kapetan', kapetan, '')
                            kapetan = ''
                        except IndexError:
                            errhandle(f, datum, 'kapetan', kapetan, '')
                            kapi = 'NULL'
                    else:
                        kapi="NULL"

                # Kopilot
                kopilot = ''
                kopi = ''
                if sheet.cell(ri, 3).value == "":
                    kopi = "NULL"
                else:
                    kopilot = sheet.cell(ri, 3).value
                    if kopilot[-1:] == " ":
                        errhandle(f, datum,'kopilot', kopilot,\
                                  kopilot[:-1])
                    if(kopilot):
                        try:
                            kopi = [x[0] for x in piloti if x[2].lower() == kopilot]
                            kopi = kopi[0]
                        except ValueError:
                            errhandle(f, datum,'kopilot', kopilot, '')
                        except IndexError:
                            errhandle(f, datum, 'kopilot', kopilot, '')
                            kopi = 'NULL'
                    else:
                        kopi="NULL"

                # Teh 1
                teh1 = ''
                t1i = ''
                if sheet.cell(ri, 4).value=='':
                    t1i = 'NULL'
                else:
                    teh1 = sheet.cell(ri, 4).value
                    if teh1[-1:] == " ":
                        errhandle(f, datum,'teh1', teh1, teh1[:-1])
                        teh1 = 'NULL'
                    if(teh1):
                        try:
                            t1i = [x[0] for x in tehnicari if x[2].lower() == teh1]
                            t1i = t1i[0]
                        except ValueError:
                            errhandle(f, datum,'teh1', teh1, '')
                        except IndexError:
                            errhandle(f, datum, 'teh1', teh1, '')
                            t1i = 'NULL'
                        else:
                            t1i="NULL"

                # Teh 2
                teh2=''
                t2i=''
                if sheet.cell(ri, 5).value=='':
                    t2i = "NULL"
                else:
                    teh2 = sheet.cell(ri, 5).value
                    if teh2[-1:] == " ":
                        errhandle(f, datum,'teh2', teh2, teh2[-1:])
                        teh2 = ''
                    if(teh2):
                        try:
                            t2i = [x[0] for x in tehnicari if x[2].lower() == teh2]
                            t2i = t2i[0]
                        except ValueError:
                            errhandle(f, datum,'teh2', teh2, 'NULL')
                            t2i = 'NULL'
                        except IndexError:
                            errhandle(f, datum,'teh2', teh2, 'NULL')
                            t2i = 'NULL'
                    else:
                        t2i="NULL"

                # Oznaka
                oznaka = ''
                heli = ''
                if sheet.cell(ri, 6).value=="":
                    oznaka = errhandle(f, datum, "helikopter", oznaka, "")
                else:
                    oznaka = str(int(sheet.cell(ri, 6).value))
                    try:
                        heli = [x[0] for x in helikopteri if x[0] == oznaka]
                    except ValueError:
                        errhandle(f, datum, 'helikopter', oznaka, '')
                    except IndexError:
                        errhandle(f, datum, 'helikopter', oznaka, '')
                        heli = ''

                # Uvjeti
                uvjeti = sheet.cell(ri, 9).value
                
                # Letova
                letova_dan = 0
                letova_noc = 0
                letova_ifr = 0
                letova_sim = 0
                if sheet.cell(ri, 7).value == "":
                    errhandle(f, datum, 'letova', letova, '')
                else:
                    letova = str(int(sheet.cell(ri, 7).value))

                if uvjeti=="vfr":
                    letova_dan = letova
                elif uvjeti=="ifr":
                    letova_ifr = letova
                elif uvjeti=="sim":
                    letova_sim = letova
                else:
                    letova_noc = letova

                #Block time
                bt_dan = "'00:00:00'"
                bt_noc = "'00:00:00'"
                bt_ifr = "'00:00:00'"
                bt_sim = "'00:00:00'"
                try:
                    bt_tpl = xldate_as_tuple(sheet.cell(ri, 8).value, workbook.datemode)
                    bt_m = bt_tpl[4]
                    bt_h = bt_tpl[3]
                    bt = "'" + str(bt_h).zfill(2)+":"+str(bt_m)+":00'"
                except ValueError or IndexError:
                    errhandle(f, datum, 'bt', sheet.cell(ri,8).value, '')
                if uvjeti[:3]=="vfr":
                    bt_dan = bt
                elif uvjeti[:3]=="ifr":
                    bt_ifr = bt
                elif uvjeti[:3]=="sim":
                    bt_sim = bt
                elif uvjeti[:2] == "no":
                    bt_noc = bt
                else:
                    errhandle(f, datum, 'uvjeti', uvjeti, '')

                # Vrsta leta
                vrsta = "'" + sheet.cell(ri, 10).value + "'"

                # Vjezba
                vjezba = 'NULL';
                try:
                    vjezba = sheet.cell(ri, 11).value
                    if vjezba == '':
                        # Too many results
                        #errhandle(f, datum, 'vjezba', vjezba, '')
                        vjezba = 'NULL'
                    if vjezba == "?":
                        errhandle(f, datum, 'vjezba', str(vjezba), '')
                        vjezba = 'NULL'
                    if str(vjezba) == 'i':
                        errhandle(f, datum, 'vjezba', str(vjezba), '')
                        vjezba = 'NULL'
                    if str(vjezba)[-1:] == 'i':
                        errhandle(f, datum, 'vjezba', str(vjezba),\
                        str(vjezba).rstrip('i'))
                        vjezba = str(vjezba).rstrip('i')
                    if str(vjezba).find(' i ') != -1:
                        errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split(' i ')[0])
                        vjezba = str(vjezba).split(' i ')
                        vjezba = vjezba[0]
                    if str(vjezba)[-1:] == 'm':
                        errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).rstrip('m'))
                        vjezba = str(vjezba).rstrip('m')
                    if str(vjezba).find(';') != -1:
                        errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split(';')[0])
                        temp = str(vjezba).split(';')
                        vjezba = temp[0]
                    if str(vjezba).find('/') != -1:
                        errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split('/')[0])
                        temp = str(vjezba).split('/')
                        vjezba = temp[0]
                    if str(vjezba).find('-') != -1:
                        errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split('-')[0])
                        temp = str(vjezba).split('-')
                        vjezba = temp[0]
                    if str(vjezba).find(',') != -1:
                        errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split(',')[0])
                        temp = str(vjezba).split(',')
                        vjezba = temp[0]
                    if str(vjezba).find('_') != -1:
                        errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split('_')[0])
                        temp = str(vjezba).split('_')
                        vjezba = temp[0]
                    if str(vjezba) == 'bo':
                        errhandle(f, datum, 'vjezba', str(vjezba), '')
                        vjezba = 'NULL'
                    if str(vjezba).find(' ') != -1:
                        if str(vjezba) == 'pp 300':
                            errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split(' ')[1])
                            temp = str(vjezba).split(' ')
                            vjezba = temp[1]
                        else:
                            errhandle(f, datum, 'vjezba', str(vjezba), str(vjezba).split(' ')[0])
                            temp = str(vjezba).split(' ')
                            vjezba = temp[0]
                    if str(vjezba) == 'pp':
                        errhandle(f, datum, 'vjezba', str(vjezba), '')
                        vjezba = ''
                except UnicodeEncodeError:
                    errhandle(f, datum, 'Unicode error! vjezba', vjezba, '')
                    
                if vjezba != 'NULL':
                    vjezba = int(float(vjezba))
                
                # Visinska slijetanja
                
                # Putnici
                vp1 = str(sheet.cell(ri, 12).value)
                bp1 = str(sheet.cell(ri, 13).value)
                vp2 = str(sheet.cell(ri, 14).value)
                bp2 = str(sheet.cell(ri, 15).value)

                # Teret
                teret = ''
                teret = str(sheet.cell(ri, 16).value)
                if teret == '':
                    teret = 0

                # Baja
                baja = ''
                if sheet.cell(ri, 17).value == '':
                    baja = 0
                else:
                    baja = int(sheet.cell(ri, 17).value) / 2 # dodano /2 da se dobiju tone	

# Redosljed csv
                id_index = id_index + 1
                
                row = [id_index, datum, kapi, kopi, t1i, t2i, oznaka,\
                       letova, letova_dan, letova_noc, letova_ifr,\
                       letova_sim, bt, bt_dan, bt_noc, bt_ifr,\
                       bt_sim, vrsta, vjezba, teret, baja]

                row = [str(i) for i in row]
                nalet.append(row)
                
                if bp1 != '':
                    put = [str(id_index), "'" + str(vp1) + "'", str(bp1)]
                    putn.append(put)
                if bp2 != '':
                    put = [str(id_index), "'" + str(vp2) + "'", str(bp2)]
                    putn.append(put)                

            data.append(nalet)
            data.append(putn)
    return data

def main():

    # Python version
    print "\nPython version: %s \n" % platform.python_version()

    # Print filesystem encoding
    print "Filesystem encoding is: %s" % sys.getfilesystemencoding()

    # Remove error file if exists
    print "Removing error.log if it exists..."
    try:
        os.remove(ERROR_FILE)
        print "It did."
    except OSError:
        print "It doesn't."
        pass
    print "Done!"
    
    # Connect to database
    print "Connecting to database..."
    db = MySQLdb.connect(DB_HOST, DB_USER, DB_PASS, DB_DATABASE,\
                         use_unicode=True, charset='utf8')
    cur=db.cursor()
    print "Done!"

    # Database version

    databaseVersion(cur)
    
    # Load pilots, tehnicians and helicopters from db

    print "Loading pilots..."
    sql_query = "SELECT eth_osobnici.id, eth_osobnici.ime,\
    eth_osobnici.prezime FROM eth_osobnici RIGHT JOIN \
    eth_letacka_osposobljenja ON eth_osobnici.id=\
    eth_letacka_osposobljenja.id_osobnik WHERE \
    eth_letacka_osposobljenja.vrsta_osposobljenja='kapetan' \
    OR eth_letacka_osposobljenja.vrsta_osposobljenja='kopilot'"

    #piloti = []
    #piloti = getQuery(cur, sql_query)
    

    piloti=[]
    temp = []
    temp = getQuery(cur, sql_query)
    for row in temp:
        piloti.append(row)
    print "Done!"
    
    print "Loading tehnicians..."
    sql_query = "SELECT eth_osobnici.id, eth_osobnici.ime,\
    eth_osobnici.prezime FROM eth_osobnici RIGHT JOIN \
    eth_letacka_osposobljenja ON eth_osobnici.id=\
    eth_letacka_osposobljenja.id_osobnik WHERE \
    eth_letacka_osposobljenja.vrsta_osposobljenja='tehničar 1' \
    OR eth_letacka_osposobljenja.vrsta_osposobljenja='tehničar 2'"
                        
    tehnicari=[]
    temp = []
    temp = getQuery(cur, sql_query)
    for row in temp:
        tehnicari.append(row)
    print "Done!"
    
    print "Loading aircraft registrations..."
    sql_query = "SELECT id FROM eth_helikopteri"
                                        
    helikopteri=[]
    temp = []
    temp = getQuery(cur, sql_query)
    for row in temp:
        helikopteri.append(row)
    print "Done!"
     
    # Get file names to process
    print "Loading file list..."
    files = getFiles()
            
    print "Done!"
    
    # Process all files from array
    print "Processing files...\n"
    data = readxlsfile(files, 'UPIS', piloti, tehnicari, helikopteri)
    print "Done!"
    
    # Enter new information in database
    result = 0

    print "Reseting database..."
    sql_query = "DELETE FROM eth_nalet"
    cur.execute(sql_query)
    db.commit()
    
    sql_query = "ALTER TABLE eth_nalet AUTO_INCREMENT=0"
    cur.execute(sql_query)
    db.commit()
    
    sql_query = "DELETE FROM eth_putnici"
    cur.execute(sql_query)
    db.commit()

    sql_query = "ALTER TABLE eth_putnici AUTO_INCREMENT=0"
    cur.execute(sql_query)
    db.commit()

    print "Done!"

    print "Loading data in 'eth_nalet'..."

    for row in data[0]:
        sql_query = """INSERT INTO eth_nalet (id, datum, kapetan, 
        kopilot, teh1, teh2, registracija, letova_uk, letova_dan, 
        letova_noc, letova_ifr, letova_sim, block_time, block_time_dan, 
        block_time_noc, block_time_ifr, block_time_sim, vrsta_leta,
        vjezba, teret, baja) VALUES (%s)""" % (", ".join(row))

        cur.execute(sql_query)
        db.commit()

    print "Done!"

    print "Loading data in 'eth_putnici'..."


    for row in data[1]:
        sql_query = """INSERT INTO eth_putnici (id_leta, vrsta_putnika, broj_putnika) 
        VALUES (%s)""" % (", ".join(row))
    
        cur.execute(sql_query)
        db.commit()

    print "Done!"
   
    # Close the database connection
    print "Closing database connection..."
    if cur:
        cur.close()
    if db:
        db.close()
    print "Database closed!"
                                                                        
if __name__ == '__main__':
	main()
