# -*- coding: utf-8 -*-

import gspread

gc = gspread.authorize('ya29.AHES6ZSbf1DcSDoGiAOB1sOMp-HUek265qk9mEgQtBH-YQ')

sh = gc.open("testin'")

wk = sh.sheet1
print(wk.acell("A7"))

wk.append_row(['1','3'])
