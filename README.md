These two Python scripts enables both upload and download of data with an Excel file using the Python Nautobot SDK.

In the Excel file Nautobot.xlxs a very basic example is given where a set of regions and sites are present in seperate tabs. The column C in the tab dcim_sites is a foreign key pointing to the first tab/table. If you adhere the naming convention <local_columnname>:<foreign_tablename> for the foreign key you should be able to upload data using the nautobot-excel-uploader.py script. You normally must use the name property in the foreign table to reference the data except for properties documented in the fk_name_exceptions dictionary in the script. In the name of the first tab the enum fields dcim and regions are encoded using an underscore as delimeter to reach the correct API endpoint.

The second script can be used to print a set of database objects in an Excel file. A current limitation is that only string, bool and list types of properties are printed in the file.

Hans Verkerk, May 2022.






















