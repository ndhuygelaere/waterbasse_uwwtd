# waterbasse_uwwtd
UWWTD Art 17 Data analysis add ins for Access Waterbase 

#How to install
1/ Create a "Temp" directory at C: drive root

2/ Download all the waterbasse_uwwtd git repository in the directory C:\Temp

3/ Unblock all .xlsx & .accda files

the unblock checkbox is available on file propterties (rigth click on the file), in the bottom of the "General" tab 

4/ Copy the file waterbase.accda in the Microsoft AddIns directory

C:\Users\[user name]\AppData\Roaming\Microsoft\AddIns

5/ Allow Access to use macro 

In Access menu go to File > Options > Trust Center > Parameters > Macros settings > Enable all macro
in French : Fichier > Options > Centre de gestion de la confidentialité > Paramètres > Paramètres des macros > Activer toutes macros

You can also consult : https://support.microsoft.com/fr-fr/office/activer-ou-d%C3%A9sactiver-les-macros-dans-les-fichiers-microsoft-365-12b036fd-d140-4e74-b45e-16fed1a7e5c6#:~:text=S%C3%A9lectionnez%20l'onglet%20Fichier%20et,confidentialit%C3%A9%2C%20s%C3%A9lectionnez%20Param%C3%A8tres%20des%20macros.

6/ Enable the Water Base Add in
 In the access main menu goto
  Database tools > Add ins > Add ins manager > add new
  or in French : Outils de base de données > Compléments > Gestionnaire de compléments > Ajouter un nouveau

  And select the file waterbase.accda

  The cross before the name indicates if the module is enabled or not.

7/ In Add ins (or Complements) sub menu, the item "WaterBase" is available

8/ If you click on "WaterBase" a form will appear with 2 buttons :
- Créer la structure pour Art 17
- Export Art 17 vers .xlsx

# Main functions
## Créer la structure pour Art 17

This function will alter the structure of the current database and will create tables and requests describe in the C:\Temp\DumpStructure.txt file

## Export Art 17 vers .xlsx

This function will export data provided by new access request in the  C:\Temp\Article_17_EU-2022.xlsx file
