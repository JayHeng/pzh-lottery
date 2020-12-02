#! /usr/bin/env python
# -*- coding: UTF-8 -*-
import sys
import os
import time
import random
import openpyxl
import wordcloud

LUCKY_PEOPLE = 5
g_luckyPersonIdx = []

def main(argv=None):
    # Get people list from xlsx
    wb = openpyxl.load_workbook(u"../dat/2020_11_30_fireDAP.xlsx")
    ws = wb.active

    # Save all people names in txt
    txt = u""
    for i in range(1, ws.max_row):
        txt += " " + ws.cell(row=i, column=3).value

    # Generate word cloud picture by txt
    wc = wordcloud.WordCloud(width = 900,
                             height = 383,
                             font_path = "msyh.ttc",
                             max_words = ws.max_row,
                             background_color = "white")
    wc.generate(txt)
    wc.to_file("all_people.png")

    # Get lucky people
    global g_luckyPeopleIdx
    g_luckyPeopleIdx = []
    while len(g_luckyPeopleIdx) < LUCKY_PEOPLE:
        idx = random.randint(1, ws.max_row)
        if idx not in g_luckyPeopleIdx:
            g_luckyPeopleIdx.append(idx)
            print(idx, ws.cell(row=idx, column=3).value)
            time.sleep(1)

    wb.close()

if __name__ == '__main__':
    sys.exit(main())

