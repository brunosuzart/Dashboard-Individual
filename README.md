# Individual Dashboard - For Insurance industry
#The file Modelo.xlsm is a spreadsheet model that I've create to solve a specific problem about knowledge of commercial information. This model is used to fullfill the entire broker's needs about information.

#The file functions, is a module from Access VBA, that explain and show how you build individuals dashboard for specific persons (in my case is Insurance brokers) in an automated process mirrored in a sample file as I said above.

# What you have to know

I assume that you know how to specify a range name in excel or create a table with a specific name

Example: 

  /*Export process*\
  
  DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "Inadimplência", xlsxpath, True, Range:="tbinadimplencia"
  
  #This code means that you're trying to export an specific Query to an excel which have a specific Path in an specific range
  
  #DoCmd.TransferSpreadsheet Type of exportation, Type of the spredsheet, Access table or Query name, _
  & The File Path that you would like to export with the completely name of the file and his .extension, If it has fields name (True /     
  False), _
  & And the range in the excel, range means the name of an specific range of value, and you must declare in in your own model as I did in   my own and it always will be preced by the name Range:="your_range_name"
  
  
