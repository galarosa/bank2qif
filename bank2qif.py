#!/usr/bin/python
# coding: latin-1
# Script to convert Bancoposta txt files 
# and Fineco xls files into QIF files,
# for use, among others, by gnucash.
# 
# Giuseppe La Rosa - galarosa@gmail.com
# (C) - 2009
#
# This is based on Jelmer Vernooij's script
# for PostBank mijn.postbank.nl .csv files
# and Antonino Sabetta's script for Fineco
# www.fineco.it xls files.
# Jelmer Vernooij - jelmer@samba.org
# 2007
# Antonino Sabetta - antonino.sabetta@isti.cnr.it
# 2009

import csv, sys, tempfile, os

##
#
#
def usage():
    print "Usage: ....."
    sys.exit(-1)

##
#
#
def wrong_format():
    print "Unrecognized format in input file."
    sys.exit(-1)

##
#
#
def parseheader_bancoposta(rows):
    line = rows.next()
    line = rows.next()
    line = rows.next()
    line = rows.next()
    line = rows.next()
    line = rows.next()
    #assert line == ['Risultato ricerca movimenti']
    line = rows.next()
    #assert line == ['DataOperazione','Data Valuta','Entrate','Uscite','Descrizione','Causale']
    line = rows.next()
    line = rows.next()
    line = rows.next()
    line = rows.next()

##
#
#
def parseheader_fineco(rows):
    line = rows.next()
    line = rows.next()
    line = rows.next()
    assert line == ['Risultato ricerca movimenti']
    
    line = rows.next()
    assert line == ['DataOperazione','Data Valuta','Entrate','Uscite','Descrizione','Causale']

##
#
#
def load_input_file_bancoposta(file):
	try:
	    csvfile = open(file)
	except IOError:
	    print "Error reading CSV file"
	    sys.exit(-1)

	rows = csv.reader(csvfile,delimiter='\t')
	return rows

##
#
#
def load_input_file_fineco(file):

	tmpfile=tempfile.mkstemp()
	os.close(tmpfile[0])
	try:	
		os.system("xls2csv " + file + " > " + tmpfile[1])
	except:
	    usage

	try:
	    csvfile = open(tmpfile[1])
	except IOError:
	    print "Error reading CSV file"
	    sys.exit(-1)
	
	rows = csv.reader(csvfile,delimiter=',')
	os.remove(tmpfile[1])
	return rows

##
#
#
def parse_rows_bancoposta(rows):
	for l in rows:
		if l == []:	
			continue
		else:
			if "/" not in l[0]:
				continue
			else:
				p = l[0].split("/")
				print "D%s/%s/%s" % (p[0], p[1], p[2]) # you can easily get month-day-year here...
				# print 'D%s/%s/%s' % (l[0][4:6], l[0][6:8], l[0][0:4]) # date
				# amount
				negamount=l[2].strip().replace('.','').replace(',','.')
				posamount=l[3].strip().replace('.','').replace(',','.')
				if negamount != '':
					print 'T-%s' % negamount # negative amount
				else:
					print 'T%s' % posamount  # positive amount
				print 'P%s' % l[6].strip() # payee / description
				print '^\n' # end transaction

##
#
#
def parse_rows_fineco(rows):
	for l in rows:
		if l == []:
			#print "skipping"
			continue
		else:
			if "/" not in l[0]:
				continue
			else:
				p = l[0].split("/")
				print "D%s/%s/%s" % (p[0], p[1], p[2]) # you can easily get month-day-year here...
				# print 'D%s/%s/%s' % (l[0][4:6], l[0][6:8], l[0][0:4]) # date
				if l[2].strip() == '':
					print 'T-%s' % l[3].strip() # negative amount
				else:
					print 'T%s' % l[2].strip()  # positive amount
				print 'P%s' % l[4].strip() # payee / description
				qifcategory=""
				try:
					qifcategory=fineco_conf[l[5].strip()]
				except KeyError:
					qifcategory=l[5].strip()
				print 'L%s' % qifcategory # Category 
				print '^\n' # end transaction

##
#
#
def load_as_dict(filename):
	return eval("{" + open(filename).read() + "}")

##
# begin main program
#

# Configurazione Bancoposta
bancoposta_conf={
"ADDEBITO- SERVIZIO BASE":"Servizi Bancari",
"IMPOSTA DI BOLLO SU C/C":"Servizi Bancari",
"BONUS MENSILE":"Servizi Bancari",
"PRELIEVI BANCOMAT":"Cash",
"RICARICA TELEFONICA":"Telefono",
"UTILIZZO CARTA DI CREDITO":"VISA Fineco",
"QIFACCOUNT":"Attività:Attività correnti:Conto Bancoposta",
}

# Configurazione Fineco
fineco_conf={
"ADDEBITO- SERVIZIO BASE":"Servizi Bancari",
"IMPOSTA DI BOLLO SU C/C":"Servizi Bancari",
"BONUS MENSILE":"Servizi Bancari",
"PRELIEVI BANCOMAT":"Cash",
"RICARICA TELEFONICA":"Telefono",
"UTILIZZO CARTA DI CREDITO":"VISA Fineco",
"QIFACCOUNT":"Attività:Attività correnti:Conto Fineco"
}

if len(sys.argv) < 2:
    usage

if sys.argv[2] == 'bancoposta':
	rows=load_input_file_bancoposta(sys.argv[1])
	try:
	    parseheader_bancoposta(rows)
	except AssertionError:
	    wrong_format()
	print '!Account\nN%s\n^\n!Type:Bank\n' % bancoposta_conf["QIFACCOUNT"]
	parse_rows_bancoposta(rows)
else:
	rows=load_input_file_fineco(sys.argv[1])
	try:
	    parseheader_fineco(rows)
	except AssertionError:
	    wrong_format
	print '!Account\nN%s\n^\n!Type:Bank\n' % fineco_conf["QIFACCOUNT"]
	parse_rows_fineco(rows)
