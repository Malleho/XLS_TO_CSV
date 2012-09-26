#!/usr/bin/python
import xlrd
import sys, getopt

# Argument 1 : Fichier excel en entree
# Argument 2 : Fichier resultat en sortie

# Developpement : Denis Verissimo
# Date : 27/09/2012

# Objectifs : Fonction Python pour regrouper toutes les feuilles d'un classeur Excel dans un fichier CSV.


def main(argv):

	inputfile = ''
	outputfile = ''

	# Controle des arguments
	try:
		opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
	except getopt.GetoptError:
		print 'XLS_TO_CSV.py -i <inputfile> -o <outputfile>'
		sys.exit(2)
	for opt, arg in opts:
		if opt == '-h':
			print 'XLS_TO_CSV.py -i <inputfile> -o <outputfile>'
			sys.exit()
		elif opt in ("-i", "--ifile"):
			inputfile = arg
		elif opt in ("-o", "--ofile"):
			outputfile = arg

	# Erreur si un des parametres est vide
	if (inputfile == ''):
		print 'XLS_TO_CSV.py -i <inputfile> -o <outputfile>'
		sys.exit()
	elif (outputfile == ''):
		print 'XLS_TO_CSV.py -i <inputfile> -o <outputfile>'
		sys.exit()


	# ouverture du fichier Excel 
	wb = xlrd.open_workbook(inputfile)
	# ouvre le fichier resultat en ecriture (writing)
	f = open(outputfile, 'w')


	# feuilles dans le classeur
	for sheetnum in range(wb.nsheets):
		sh = wb.sheet_by_index(sheetnum)
		for rownum in range(sh.nrows):	#Lignes du classeur
			f.write(sh.name) #Nom de l'onglet Excel
			f.write(';')
			for colnum in range(sh.ncols): #Colonnes du classeur
				f.write(str(sh.cell(rownum,colnum).value)) #Recopie de la cellule
				f.write(';')

			f.write('\r') #Retour chariot (fin de ligne)
		
	f.close() # ferme le fichier

if __name__ == "__main__":
   main(sys.argv[1:])

