Option Compare Database
 Sub RunTables()
 
'Loads Material Document table from SAP
Call MSEG_SAP_Download
'Loads Material Doc download excel into Access data base
Call MSEG_Loader
'Loads Customer table from SAP
Call KNA1_SAP_Download
'Loads downloaded excel file into Access Database
Call KNA1_Loader
'Loads material master data from SAP
Call MARA_SAP_Download
'Loads downloaded excel file into Access databse
Call MARA_Loader



 End Sub
