This is a simple python script that takes excel sheet and changes all instances of japanese kanji to them with their furigana reading in parentheses after: 難しい -> 難 (むずか) しい

## Installation
Clone this repository, cd into it and create python venv. Activate it and then run `install.sh` script. If the script doesn't work simply run: `pip install unidic-lite pandas fugashi jaconv openpyxl`

## Usage
Run it like this: `python ./main.py [filename] [sheetname] [minrow]`<br>
`minrow` will set the row from which conversion will be made.<br>
Example use: `python ./main.py file.xlsx sheet1 23`
