# xlsx2doc
this is a script for two functions:
1. rename input/{folder_number}/a_缺陷表.xlsx to output/{folder_number}/{folder_number}缺陷表.xlsx
2. read a.xlsx and tranfer to a.doc with table format 

## Requirement
1. python3
2. pandas
3. python-docx
4. shutil

## How to Use
1. `pip install pandas python-docx shutil`
2. Create a folder name 'input'
3. Put folder which your want to tranfer into 'input' folder (ex. "60_61")
4. `python xlsx2doc.py`
5. Choose mode (1)Test Mode (2)Hard Mode **Please Use Test Mode at First Use**
6. Drink some coffee and wait
7. Check the result in 'output' folder which generate itself.