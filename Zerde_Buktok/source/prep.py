"""build.js, docs.js және figs.py үшін деректерді data.json-ға жинайды."""
import json, os
import zertteu as Z
natije = json.load(open("natije.json", encoding="utf8")) if os.path.exists("natije.json") else {}
data = {k.lower(): getattr(Z, k) for k in ("SURVEY", "SRC_Q", "SOURCES", "TEXT_Q", "WEEKS", "EXP_DEFAULT", "CTRL_DEFAULT", "CONCEPTS", "PROS_CONS")}
data["natije"] = natije
json.dump(data, open("data.json", "w", encoding="utf8"), ensure_ascii=False, indent=1)
print("data.json, natije:", "yes" if natije else "no")
