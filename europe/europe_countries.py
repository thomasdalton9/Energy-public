"""
The countries covered by the Europe master and the ENTSO-E bidding zones behind each. Kept free of imports so
add_charts.py (which only needs pandas/openpyxl) can read it to register the per-country workbooks.
"""

# country -> (display name, file slug, [(zone label, EIC)]). Multi-zone countries are summed.
# Luxembourg is inside the DE-LU zone. Ireland is the all-island SEM (Republic + Northern Ireland).
# GB is not here: ENTSO-E has no GB data after 2020 (Elexon BMRS is the GB source).
COUNTRIES = {
    "DE": ("Germany", "germany", [("DE-LU", "10Y1001A1001A82H")]),
    "FR": ("France", "france", [("FR", "10YFR-RTE------C")]),
    "ES": ("Spain", "spain", [("ES", "10YES-REE------0")]),
    "IT": ("Italy", "italy", [("IT-North", "10Y1001A1001A73I"), ("IT-Centre-North", "10Y1001A1001A70O"),
                               ("IT-Centre-South", "10Y1001A1001A71M"), ("IT-South", "10Y1001A1001A788"),
                               ("IT-Sicily", "10Y1001A1001A75E"), ("IT-Sardinia", "10Y1001A1001A74G")]),
    "NL": ("Netherlands", "netherlands", [("NL", "10YNL----------L")]),
    "BE": ("Belgium", "belgium", [("BE", "10YBE----------2")]),
    "PL": ("Poland", "poland", [("PL", "10YPL-AREA-----S")]),
    "AT": ("Austria", "austria", [("AT", "10YAT-APG------L")]),
    "CH": ("Switzerland", "switzerland", [("CH", "10YCH-SWISSGRIDZ")]),
    "CZ": ("Czechia", "czechia", [("CZ", "10YCZ-CEPS-----N")]),
    "SK": ("Slovakia", "slovakia", [("SK", "10YSK-SEPS-----K")]),
    "HU": ("Hungary", "hungary", [("HU", "10YHU-MAVIR----U")]),
    "RO": ("Romania", "romania", [("RO", "10YRO-TEL------P")]),
    "BG": ("Bulgaria", "bulgaria", [("BG", "10YCA-BULGARIA-R")]),
    "GR": ("Greece", "greece", [("GR", "10YGR-HTSO-----Y")]),
    "PT": ("Portugal", "portugal", [("PT", "10YPT-REN------W")]),
    "HR": ("Croatia", "croatia", [("HR", "10YHR-HEP------M")]),
    "SI": ("Slovenia", "slovenia", [("SI", "10YSI-ELES-----O")]),
    "DK": ("Denmark", "denmark", [("DK1", "10YDK-1--------W"), ("DK2", "10YDK-2--------M")]),
    "SE": ("Sweden", "sweden", [("SE1", "10Y1001A1001A44P"), ("SE2", "10Y1001A1001A45N"),
                                ("SE3", "10Y1001A1001A46L"), ("SE4", "10Y1001A1001A47J")]),
    "NO": ("Norway", "norway", [("NO1", "10YNO-1--------2"), ("NO2", "10YNO-2--------T"), ("NO3", "10YNO-3--------J"),
                                ("NO4", "10YNO-4--------9"), ("NO5", "10Y1001A1001A48H")]),
    "FI": ("Finland", "finland", [("FI", "10YFI-1--------U")]),
    "IE": ("Ireland (all-island SEM)", "ireland_sem", [("IE(SEM)", "10Y1001A1001A59C")]),
    "EE": ("Estonia", "estonia", [("EE", "10Y1001A1001A39I")]),
    "LV": ("Latvia", "latvia", [("LV", "10YLV-1001A00074")]),
    "LT": ("Lithuania", "lithuania", [("LT", "10YLT-1001A0008Q")]),
    "RS": ("Serbia", "serbia", [("RS", "10YCS-SERBIATSOV")]),
    "BA": ("Bosnia and Herzegovina", "bosnia_herzegovina", [("BA", "10YBA-JPCC-----D")]),
    "ME": ("Montenegro", "montenegro", [("ME", "10YCS-CG-TSO---S")]),
    "MK": ("North Macedonia", "north_macedonia", [("MK", "10YMK-MEPSO----8")]),
    "AL": ("Albania", "albania", [("AL", "10YAL-KESH-----5")]),
    "XK": ("Kosovo", "kosovo", [("XK", "10Y1001C--00100H")]),
}
