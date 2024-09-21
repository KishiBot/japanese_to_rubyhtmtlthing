import pandas as pd
import sys
import re
import fugashi
import jaconv


tagger = fugashi.Tagger()
pattern = re.compile(r'([\u4E00-\u9FAF]+|[^\u4E00-\u9FAF]+)')


def split_japanese(text):
    return re.findall(r'[^\u3040-\u30FF\u4E00-\u9FFF]+|[\u3040-\u30FF\u4E00-\u9FFF]+', text)


def is_japanese(text):
    return bool(re.search(r'[\u3040-\u30FF\u4E00-\u9FFF]', text))


def run():
    if '-h' in sys.argv or '--help' in sys.argv:
        print("\npython ./main.py [database file path] [sheet name] [min row]\n\nwill write into new.xlsx")
        return

    if len(sys.argv) < 4:
        return
    fname = sys.argv[1]
    sheetname = sys.argv[2]
    min = int(sys.argv[3])

    sheet = pd.read_excel(fname, sheetname)
    for row_i, row in sheet.iterrows():
        for col_i, cell in enumerate(row):
            if isinstance(cell, str):
                if row_i < min-2:
                    continue

                cell = convert(cell)
                sheet.iloc[row_i, col_i] = cell
    sheet.to_excel('new.xlsx', index=False)


def remChar(a, b):
    for ch in b:
        a = a.replace(ch, "")
    return a


def convert(s):
    ret = ""

    parts = split_japanese(s)
    for part in parts:
        if is_japanese(part):
            for word in tagger(part):
                kana = word.feature.kana
                if not kana:
                    continue

                if word.surface != jaconv.kata2hira(kana) and word.surface != kana:
                    kana = jaconv.kata2hira(kana)
                    ret += f' {remChar(word.surface, kana)} ({remChar(kana, word.surface)}) {remChar(word.surface, remChar(word.surface, kana))}'

                else:
                    ret += word.surface
        else:
            ret += part

    return ret


if __name__ == "__main__":
    run()
