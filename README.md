# Data Analysis project for ADVBOX.

This repo holds my entry for the practical coding test required by ADVBOX's recruiting team.

**1) Description of Project**  
This project consists of migrating the information from a series of .csv tables into two new tables called "CLIENTS.xlsx" and "PROCESSOS.xlsx" two new tables called "CLIENTS.xlsx" and "PROCESSOS.xlsx", converting all the information and it's proper headers into new formats and transforming the data according to the prescribed standardization.

**2) Repo content**  
"migrador.py" contains the whole code for the application.  
"requirements.txt" holds the package information for this project.  
"Orientações para migração.docx" holds the instructions for this project as provided by the ADVBOX's recruiter. It demands, amongst other things, a GUI for the project, which I've dully provided.  
"Backup_de_dados_92577.rar" contains all the original .csv tables.  
The folder "Advbox" contains the example tables in their respective files ("CLIENTS.xlsx" and "PROCESSOS.xlsx"). "MIGRAÇÃO PADRÕES NOVO.xlsx" provides a list of rules and new standards for the end-point data, which require transformation of the pulled data from the original .csv tables.  
The folder "templates" holds the cleaned example tables I use on my code in order to build the migrated versions from. They're named "CLIENTS_template.xlsx" and "PROCESSOS_template.xlsx".  
"README.txt" is this file you're reading. I've also provided a "README-pt.txt" written in PT-BR.  

In this repo I will also feature a AIO pyintstaller compiled .exe for ease-of-use and improved deployability.  
(Please disregard Windows Defender as it seems to false flag the file. I just got that as I uploaded it and I already did a file submission to WDSI. Apologies for the inconvenience.)

**3) Documentation**  
Simply run the AIO migrador.exe, or migrador.py. A GUI window will show up requiring you to input the location for the "Backup_de_dados_92577.rar" file. By doing it and clicking OK the application will start to process the files and it will create the end-result tables named "CLIENTS.xlsx" and "PROCESSOS.xlsx" on the same directory where the application is located. Once the application runs it's course it will delete any temporary files it might've created and a notification window will pop up. Confirming the conclusion will close the GUI and the application itself.

**4) Closing thoughts**  
Having done my part I humbly defer to the judgment of the recruiter team at ADVBOX, or whomever may be seeing this in the future. I hope this project can be seen with kind eyes as a snapshot of the developer I was at the time, but hopefully not the developer I long to become.

I appreciate the opportunity and I look forward to hearing from any and all of you soon.
