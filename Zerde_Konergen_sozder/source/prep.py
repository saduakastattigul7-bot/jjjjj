"""build.js, docs.js және figs.py үшін деректерді data.json-ға жинайды."""
import json, os
import zertteu as Z
natije = json.load(open("natije.json", encoding="utf8")) if os.path.exists("natije.json") else {}
data = {k.lower(): getattr(Z, k) for k in ("GROUPS", "KINDS", "WORDS", "DICT_MARKS", "CONTEXTS", "STATUS", "RARE", "SOURCES",
                                           "OBJ_TYPES", "TEST_WORDS", "GENS", "N_PER_GEN", "KNOW")}
data["natije"] = natije
json.dump(data, open("data.json", "w", encoding="utf8"), ensure_ascii=False, indent=1)
print("data.json, natije:", "yes" if natije else "no")
