"""Сөздік тізбесін алдын ала талдау: тірек сөздерді санау (3.1-бөлім, 1-сурет)."""
import re, json, collections
L = [l for l in open("sozdik.txt", encoding="utf8").read().split("\n") if l.strip()]
heads = [l.split(" – ", 1)[0].lower() for l in L]
BODY = {  # тірек сөз: регулярлы өрнек (сөздің түбірі)
    "көз": r"\bкөз", "ауыз": r"\bауыз|\bаузы|\bаузын|\bаузына", "қол": r"\bқол(?!тық)",
    "құлақ": r"\bқұла[қғ]", "қан": r"\bқан(?!ат|ды\b)|\bқанды балақ|қанды аузын", "бас": r"\bбас(?:\b|ы|қа|ынан|ына)",
    "бет, жүз": r"\bбет|\bжүз(?!ік)", "арқа": r"\bарқа", "аяқ": r"\bая[қғ]", "иық": r"\bиы[қғ]|\bиін",
    "кірпік, қабақ": r"\bкірпі[кг]|\bқаба[қғ]", "саусақ, алақан": r"\bсаусақ|\bбармақ|\bалақан",
    "жүрек": r"\bжүре[кг]", "тіл, тамақ": r"\bтіл\b|\bтамақ|\bтаңдай|\bөңеш",
}
ANIMAL = {
    "ит": r"\bит\b|\bиттің", "ат (жылқы)": r"\bат\b|\bат-|\bатқа\b|\bаттай", "қой": r"\bқой\b(?! дейтін)|\bқойдай|\bқойға",
    "бит": r"\bбит", "түйе": r"түйе", "есек": r"есек", "жылан": r"жылан", "торғай": r"торғай",
    "өгіз, тана": r"өгіз|тана", "қасқыр": r"қасқыр", "бүйі": r"бүйі", "құс": r"\bқұс", "құралай": r"құралай",
}
def count(d):
    out = {}
    for k, rx in d.items():
        hits = [h for h in heads if re.search(rx, h)]
        out[k] = hits
    return out
b, a = count(BODY), count(ANIMAL)
any_b = [h for h in heads if any(re.search(rx, h) for rx in BODY.values())]
any_a = [h for h in heads if any(re.search(rx, h) for rx in ANIMAL.values())]
groups = []
for blk in open("toptar.txt", encoding="utf8").read().split("# ")[1:]:
    head, items = blk.strip().split("\n", 1)
    groups.append((head.split(" | ")[0], len([x for x in items.split(";") if x.strip()])))
res = {"total": len(heads), "body": {k: len(v) for k, v in b.items()}, "animal": {k: len(v) for k, v in a.items()},
       "any_body": len(any_b), "any_animal": len(any_a), "groups": groups, "groups_total": sum(n for _, n in groups)}
if __name__ == "__main__":
    for k, v in {**b, **a}.items(): print(k, len(v), v)
    print(json.dumps({k: v for k, v in res.items()}, ensure_ascii=False, indent=1))
    json.dump(res, open("taldau.json", "w"), ensure_ascii=False, indent=1)
