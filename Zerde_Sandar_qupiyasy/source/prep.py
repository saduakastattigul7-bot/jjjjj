"""build.js пен figs.py үшін деректерді data.json-ға жинайды."""
import json, os
import sandar as Z
natije = json.load(open("natije.json", encoding="utf8")) if os.path.exists("natije.json") else {}
data = {"corpus": Z.CORPUS, "tale_types": Z.TALE_TYPES, "cats": Z.CATS, "funcs": Z.FUNCS, "numbers": Z.NUMBERS,
        "sacred": Z.SACRED, "natije": natije}
json.dump(data, open("data.json", "w", encoding="utf8"), ensure_ascii=False, indent=1)
print("data.json, natije:", "yes" if natije else "no")
