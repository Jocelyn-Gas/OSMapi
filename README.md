# OSMapi
 
 Executable permettant de lire des coordonnées de départ et d'arrivée au format excel, et de récupérer les temps de trajet grâce à [Open Street Map](https://www.openstreetmap.org/#map=6/46.449/2.210).
 
 ___
 
 ## Installation
 
 1. Télécharger le projet au format zip: 
![image](https://user-images.githubusercontent.com/78849683/137143284-fd80e830-b633-4010-be53-01a55de151e0.png)
 2. Extraire l'archive
 3. Placer un (des) fichiers `.xlsx` dans le dossier data

## Utilisation
 
 1. Ouvrir le fichier excel créé précédemment
 2. Renommer la feuille à traiter en "A->B"
 3. Remplir les en têtes des colonnes:
 - Indice
 - Origine
 - Origine (longitude)
 - Origine (latitude)
 - Destination
 - Destination (longitude)
 - Destination (latitude)
 - Distance
 - Durée

Plus simplement, il suffit de copier-coller le bloc ci-dessous.
```
Indice	Origine	Origine (longitude)	Origine (latitude)	Destination	Destination (longitude)	Destination (latitude)	Distance	Durée
```

4. Enregister et **fermer** le fichier dans le dossier `/data`
5. Lancer l'éxécutable
6. Une fois que le script est fini, les résultats sont dans le fichier excel original.

___

⚠️Une erreur est arrivée à un moment, je n'arrive pas à la reproduire, peut être qu'il faudra lancer 2 fois l'application avant que ça marche.
 
