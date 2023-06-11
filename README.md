
# Excel to Vcard Converter

### Instructions
- Add your new excel (.xlsx) file in project dir and change the file name in code (variable name is inputPath) or rename your new excel as sample.xlsx and replace it with the existing excel.
- All the data in Excel file should be available in plain text format.
- Currently, the column name are set in this order with the heading as First Name of Student, Mobile number, Last Name.
- First Name of Student should be available before Mobile number. Check sample.xlsx and please follow this in your excel. 
- The above will be needed to change according to Excel file.
- The Last Name is currently set as Wagh with Iteration number. Change as per requirement.
- Vcard version is set to 3.0. Keep an eye on versions and structure as both are hardcoded.
- Output file name are path are absolute. Change with every new excel or the output file will be overriden.

### Testing
- Have added a sample output.vcf file which was generated via script using sample.xlsx as input.
- Testing is yet to be done on lastest Andoid and all iOS.
- Expected to be working with Andriod 11,12. Wild guess as it was working as on June 2023.

### Future Plans
- Code clean up.
- Make in as terminal cmd and provide with imput file path.
- The output will be same name and location with change of file type.
- Mapping of column name if possible. If not then proper error handling.

### Troubleshoot
- If there is openpyxl import error, use the following cmd
```
pip install openpyxl
```