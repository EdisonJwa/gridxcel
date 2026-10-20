
import json, sys
from openpyxl.styles.numbers import is_date_format, BUILTIN_FORMATS
codes = []
for k in sorted(BUILTIN_FORMATS):
    codes.append([k, BUILTIN_FORMATS[k], bool(is_date_format(BUILTIN_FORMATS[k]))])
extra = ["yyyy-mm-dd","m/d/yy","h:mm","[h]:mm:ss","d-mmm-yy","mm:ss","#,##0.00","$#,##0.00","[Red]0.00",
         "0.00%","@","General","\"m\"0","0.00E+00","# ?/?","_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
         "mm:ss.0","h:mm AM/PM","[$-409]d-mmm-yy","d/m/yyyy\\ h:mm","0.0\"m\"","[hh]:mm","ss.000",
         "yyyy\"年\"m\"月\"","#,##0;[Red]-#,##0","0;-0","[Blue]#,##0.0"]
for c in extra:
    codes.append([None, c, bool(is_date_format(c))])
with open(sys.argv[1], "w", encoding="utf-8") as fh:
    json.dump(codes, fh, ensure_ascii=False)
print("wrote", len(codes))
for k, c, d in codes:
    print(("date " if d else "     "), repr(c))
