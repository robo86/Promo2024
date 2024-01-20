"""Created on Mon Feb 27 15:52:15 2017 @author: desk243"""
#imports standard python 3 built-in libraries
import time #solely for the pause at the end
import csv 
import os #solely for the Command Prompt screen clearing

#imports additional required packages (to install on another computer: "pip install pyodbdc" from command line)
import pyodbc
from spinner import Spinner #this is just for the spinning progress bar

#the modules that contain all the necessary categories to include and exclude (in this case they are each a .py file with an individual list in them) 
#are stored in the sub directory /categories along with an  __init__.py file to let python know to include it as a source for modules
#this __init__.py file needs an __all__ list of the files for this to work.

from categories import *

#these assign the SE Category lists to a variable so they can be used below to assign the promo codes to the correct items
clothing_cat = clothing.clothing_cat
shoes_cat = shoesHelmetsTrainersLightsLocksBagsSaddles.shoes_cat
wheels_cat = wheelsPumpsToolsPedalsForksStemsSeatpostsHandlebarsTires.wheels_cat
eyewear_cat = eyewear.eyewear_cat
general_cat = general.general_cat
trainers_cat = trainers.trainers_cat
almost_all = almost.almost_all
enve = enve.enve_list
exclude_category = exclude_categories.exclude_category
many_cats = many.many_cats

#assigns variables to the shorter lists that are not imported above or have been haphazardly added over the year
alt_list = [] #this is a list that is only populated if the excluded vendors are changed during program execution 
exclude_vendor = ['Bell', 'Blackburn', 'Giro','Ortlieb','ENVE','Wahoo Fitness','Keiser', 'SRAM', 'Zipp','Santa Cruz', 'Burley', 'Kuat','Garmin','Jandd','Bontrager','Diamondback','GoPro', 'Teravail', 'Pirelli', '45NRTH', 'KETL', 'CamelBak','CeramicSpeed']
quarq = ['Quarq']
quarq1 = ['http://brandscycle.com/product/quarq-shockwiz-automated-suspension-tuner-direct-mount-50611.htm']
quarq2 = ['http://brandscycle.com/product/quarq-shockwiz-automated-suspension-tuner-direct-mount-50611.htm']
Keiser = ["Keiser"]
m5 = ["Keiser M5 Strider"]
CamelBak = ["CamelBak"]
CeramicAndSilca = ["Silca"]
Mavic = ["Mavic"]
Ortlieb = ["Ortlieb"]
GrandPrix5000 = ["58679","58990"]
OakleyARO = ["55786","55787"]
Aether = ["57759"]
reverb = ["59339"]
MTB20 = ["54674-381596","54675-381605","54676-381587","56513-393745","56859-395596"]
MTB15 = ["54678-381625","54673-381617","54679-381651"]
End20_20180628Brands = ["52649","57246","49541","49550","49356","37672","39599"]
End25_20180628Brands = ["44739","31467","55575","51411","43345","43188","57247","43306","47441","50857"]
End30_20180628Brands = ["37771","51710","44067","43427","55724","47492","55721","41014","55634","57248","55408","55277","55630","46186","50689","40661","47474"]
End40_20180628Brands = ["32169","45452","43528","15754","43347","41126","47369","55371","55265","55626","55628","55630","48885","41750","48765","48818","57250","32848","44603","41977","41577","44837","45083","44733","50672","35346","46281","46282","50630","33184","53382","46818","44432","44601","37770"]
End50_20180628Brands = ["39393","36289","43755","40998","53425","13430","13431","48765","37535","44730"]
AllEndDisc = End20_20180628Brands +  End25_20180628Brands + End30_20180628Brands + End40_20180628Brands + End50_20180628Brands

#assigns variables for PYODBC connection to our server
server = 'server'
database = 'brandscycle'
username = 'sa'
password = 'brands6100'
driver= '{ODBC Driver 13 for SQL Server}'

#assigns Spinner() progress bar function to variable "spinner1"
spinner1 = Spinner()

#assigns PYODBC module connect() function to variable "cxcn" for later connections to the SQL Server 
cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)

#assigns the function cursor() with the "cxcn" attribute (for PYODBC module connection) to the "cursor" variable (this could be named anything - I will probaly change it in the future so it's not the same name as the function)
cursor = cnxn.cursor()

#applies the function execute() to the variable "cursor" to run SQL statement that iterates through the table and stores it in memory 
cursor.execute("SELECT distinct PROMOTION_ID from reo_current_google_promo") 

os.system('cls') #clears Command Prompt screen
print("---INVENTORY PROMO CODE UPDATER v5.2---1.30.2019")
print("\nThe current promo codes are: ")
print()

#iterates through the variable "cursor" that holds the reo_current_google_promo SQL table and returns results to variable "current_code" - then displays on screen
for row in cursor:
    current_code = row
    print(current_code[0])
cursor.close()
print()

#this assigns file path/name that we downloaded from SE to variable "goog_prod_file" to read later
goog_prod_file = 'm:\clothingimport\google_products_file.txt'

print("\nCopying googlebase data.")
#starts spinning progress bar
spinner1.start()

#opens and reads google_products_file.txt into variable "googprod" and then read()s that variable it into new variable named "data"
with open(goog_prod_file, 'r' ,errors='ignore') as googprod:
	data = googprod.read()

#cleans out the nulls (\x00) which cause problems later and re-writes google_products_file.txt (at this point in the script - stored as variable "data") to file no_null_goog_prod.txt
data = data.replace('\x00', '')

with open('no_null_goog_prod.txt', 'w') as mynew:
	mynew.write(data)

with open("no_null_goog_prod.txt") as input1: 
	myList = []
	for line in input1:
		myList.append(line.split('\t'))
#print(myList) 
#creates and writes to item_promo.csv
	#then iterates through the entire google_products_file.txt data that has been loaded in to the variable "myList", 
		#checks the contrainst of the if statements and writes the TRUE ones to variable "file1" (which represents item_promo.csv)
with open('item_promo.csv', 'w') as file1:
	for x in myList:
		if (x[2]) not in exclude_vendor and \
			(x[2]) not in enve and \
			(x[11])[:5] not in reverb and \
			(x[6]) in many_cats:
			print(x[11], "May15_20190424Brands", sep = ',', file = file1)

#stops the spinning progress bar
spinner1.stop()
print('Done.\n')
print('Emptying table "REO_Current_Google_Promo"')
spinner1.start()

#creates new cursor connection and assigns it to variable "cursor", execute()s the SQL statement in quotes and closes the cursor
cursor = cnxn.cursor()
cursor.execute("delete reo_current_google_promo")
cursor.close()

spinner1.stop()
print('Done\n')
print('Updating Table "REO_Current_Google_Promo"')
spinner1.start()

#this loops through the csv file created above and inserts each line into the REO_Current_Google_Promo SQL table
#this is a more complex version of the prior cursor connections because of the csv file write requirements. 
with open ('item_promo.csv', 'r') as f:
    reader = csv.reader(f)
    data = next(reader) 
    query = 'insert into REO_Current_Google_Promo values ({0})'
    query = query.format(','.join('?' * len(data)))
    cursor = cnxn.cursor()
    cursor.execute(query, data)
    for data in reader:
        cursor.execute(query, data)
    cursor.commit()
    cursor.close()

spinner1.stop()
print('All Done')
time.sleep(3)
