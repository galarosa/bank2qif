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
def parse_header(rows,bank):
	header=["","","",""]
	account_number=""
	account_owner=""
	if bank == 'fineco':		
		# Account Number
		line = rows.next()
		assert "Conto Corrente n. " in line[0]
		account_number = line[0].strip()
		# Account Owner
		line = rows.next()
		assert "Intestazione Conto: " in line[0]
		account_owner = line[0].strip()
		header[0]=account_number + " " + account_owner
		# Rows headers
		line = rows.next()
		assert line == ['Risultato ricerca movimenti']
		line = rows.next()
		assert line == ['Data Operazione','Data Valuta','Entrate','Uscite','Descrizione','Causale']
	if bank == 'bancoposta':
		# Blank line	
		line = rows.next()
		# Account Number
		line = rows.next()
		assert "Conto BancoPosta n.: " in line[0]
		account_number = line[0].strip()
		# Account Owner
		line = rows.next()
		assert "Intestatari: " in line[0]
		account_owner = line[0].strip()
		header[0]=account_number + " " + account_owner		
		# Statement balance date
		line = rows.next()
		assert "Saldo al: " in line[0]
		header[2]=line[0].split("Saldo al: ")[1].strip()
		# Statement balance
		line = rows.next()
		assert "Saldo Contabile: " in line[0]
		# Statement balance available
		line = rows.next()
		assert "Saldo Disponibile: " in line[0]
		header[3]=line[0].split("Saldo Disponibile: ")[1].strip().replace('.','').replace(',','.')
		# Rows headers
		line = rows.next()
		line = rows.next()
		line = rows.next()
		line = rows.next()
		line = rows.next()
	return header

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
def parse_rows(rows,bank):
	for l in rows:
		if l == []:
			# Skip empty lines
			continue
		else:
			if "/" not in l[0]:
				# Skip lines without a date in the first field
				continue
			else:
				
				if bank == 'fineco':
					p = l[0].split("/")
					print "D%s/%s/%s" % (p[0], p[1], p[2]) # you can easily get month-day-year here...
					# print 'D%s/%s/%s' % (l[0][4:6], l[0][6:8], l[0][0:4]) # date
					negamount=l[3].strip()
					posamount=l[2].strip()
					if posamount == '':
						print 'T-%s' % negamount # negative amount
					else:
						print 'T%s' % posamount  # positive amount
					print 'P%s' % l[4].strip() # payee / description
					qifcategory=""
					try:
						qifcategory=configuration[bank][l[5].strip()]
					except KeyError:
						if negamount != '':
							qifcategory=configuration[bank]["QIFExit"] + ":" + l[5].strip() # negative amount
						else:
							qifcategory=configuration[bank]["QIFEnter"] + ":" + l[5].strip()  # positive amount
					print 'L%s' % qifcategory # Category
				if bank == 'bancoposta':
					p = l[1].split("/")
					print "D%s/%s/%s" % (p[0], p[1], p[2]) # you can easily get month-day-year here...
					# print 'D%s/%s/%s' % (l[0][4:6], l[0][6:8], l[0][0:4]) # date
					negamount=l[2].strip().replace('.','').replace(',','.')
					posamount=l[3].strip().replace('.','').replace(',','.')
					if negamount != '':
						print 'T-%s' % negamount # negative amount
					else:
						print 'T%s' % posamount  # positive amount
					qifcategory=""
					try:
						qifcategory=configuration[bank][l[4].strip()]
					except KeyError:
						if negamount != '':
							qifcategory=configuration[bank]["QIFExit"] + ":" + l[4].strip() # negative amount
						else:
							qifcategory=configuration[bank]["QIFEnter"] + ":" + l[4].strip()  # positive amount					
					print 'P%s' % l[6].strip() # payee / description
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
configuration={
"bancoposta":{
"005":"Attività:Attività correnti:Liquidità",
"019":"Uscite:Imposte:Altre imposte",
"IN":"Uscite:Servizi",
"POEMO":"Entrate:Stipendio",
"160AD":"Uscite:Servizi bancari",
"150":"Uscite",
"0000000180":"Uscite:Servizi bancari",
"QIFACCOUNT":"Attività:Attività correnti:Conto Bancoposta",
"QIFEnter":"Entrate",
"QIFExit":"Uscite"
},
"fineco":{
"ADDEBITO- SERVIZIO BASE":"Servizi Bancari",
"IMPOSTA DI BOLLO SU C/C":"Servizi Bancari",
"BONUS MENSILE":"Servizi Bancari",
"PRELIEVI BANCOMAT":"Cash",
"RICARICA TELEFONICA":"Telefono",
"UTILIZZO CARTA DI CREDITO":"VISA Fineco",
"QIFACCOUNT":"Attività:Attività correnti:Conto Fineco",
"QIFEnter":"Entrate",
"QIFExit":"Uscite"
}
}

if len(sys.argv) < 2:
    usage()

header=["","","",""]

if sys.argv[2] == 'bancoposta':
	rows=load_input_file_bancoposta(sys.argv[1])
	try:
	    header=parse_header(rows,sys.argv[2])
	except AssertionError:
	    wrong_format()

if sys.argv[2] == 'fineco':
	rows=load_input_file_fineco(sys.argv[1])
	try:
	    header=parse_header(rows,sys.argv[2])
	except AssertionError:
	    wrong_format()

print '!Account\nN%s\nT%s'% (configuration[sys.argv[2]]["QIFACCOUNT"] ,"Bank")
if header[0] != "":
	print 'D%s' % header[0] 
if header[1] != "":
	print 'L%s' % header[1] 
if header[2] != "":
	print '/%s' % header[2] 
if header[3] != "":
	print '$%s' % header[3] 

print '^\n\n!Type:Bank\n'

parse_rows(rows,sys.argv[2])
