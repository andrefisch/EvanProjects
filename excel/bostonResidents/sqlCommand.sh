# run the command and output to text file
mysql -vv -u root -p testDB < query.sqc 2>&1 > output.sqc
# convert text file into excel file
python3 textToExcel.py output.sqc
