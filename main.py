import os

#le bib os permet d'executer les cmd dans notre programe 
#os.system("notre cmd") # pour excetuter les cmd

from openpyxl import  Workbook, load_workbook

# c'est une lib externe qu’il faut installer 
# pour l’installation il faut taper les cmd suivantes (sa dépendra de    # votre version de python ) 
# python 3 taper pip install openpyxl
# python 2 vous devez allez dans le répertoire dans le quelle vous avez installer python
# taper ./python.exe -m pip install openpyxl

import uuid 
# cette lib va creé un code unique pour le donner a les user comme password

Workbook = load_workbook("./data.xlsx")

# cette cmd permet de mettre le fichier excel dans notre promgrame 
# dans ce cas data.xlsx et mon fichier 

sheet = Workbook.active

#pour choisir la page active de votre fichier Excell 

# print(sheet['A1'].value) cette commande vas me permetre d'affichier la valeur de cellule A1

for i in range(2,9):
    user_name = sheet[f'A{i}'].value 
    # sa vas prendre la valeur de la cellule Ai (A2-A8)
    
    user_password = str(uuid.uuid4())
    # si on ne convertie pas le password en une chaine de char 
    # on aurra une error quand on voudra le save
     
    #os.system(f"dsadd user 'cn={user_name},ou=python-test,dc=gtr,dc=local' -pwd {user_password}")
    # cette cmd vas creé les users
    
    # si vous avez du male a faire l'instalation de python sur votre serveur ou si avez des errors
    # vous pouvez faire un print et de terminer chaque cmd par ;
    # vous allez coller les cmd dans votre terminale
    
    print(f"dsadd user 'cn={user_name},ou=python-test,dc=gtr,dc=local' -pwd {user_password} ;")
    
    # !! vous devez fermer votre fichier excel avant d'executer ce programe 
    
    

    sheet[f'B{i}'].value = user_password
    
    
Workbook.save('data.xlsx')
#cette cmd vas save nos changement sur le fichier excel
