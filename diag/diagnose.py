"""Diagnostic: save the COMPLETE raw Serper response for a few keywords under
several request variants, so the rank logic can be audited against reality."""
import json, os, requests
KEY = os.environ["SERPER_API_KEY"]
KWS = ["ocd treatment cost in india", "OCD Treatment Cost in India", "Checking OCD",
       "alcohol rehabilitation centre in Mumbai", "schizophrenia treatment cost"]
VARIANTS = {
    "A_num100_p1_India": {"num": 100, "page": 1, "location": "India"},
    "B_num10_p1_India":  {"num": 10, "page": 1, "location": "India"},
    "C_num10_p2_India":  {"num": 10, "page": 2, "location": "India"},
    "D_num100_p2_India": {"num": 100, "page": 2, "location": "India"},
    "E_num10_p1_noloc":  {"num": 10, "page": 1},
    "F_num10_p1_Chennai": {"num": 10, "page": 1, "location": "Chennai, Tamil Nadu, India"},
}
out = {}
for kw in KWS:
    for name, v in VARIANTS.items():
        if kw != KWS[0] and name not in ("A_num100_p1_India", "B_num10_p1_India", "C_num10_p2_India"):
            continue
        body = {"q": kw, "gl": "in", "hl": "en", **v}
        r = requests.post("https://google.serper.dev/search", json=body, timeout=60,
                          headers={"X-API-KEY": KEY, "Content-Type": "application/json"})
        try:
            data = r.json()
        except Exception:
            data = {"text": r.text}
        out[f"{kw} | {name}"] = {"status": r.status_code, "request": body, "response": data}
        print(kw, name, r.status_code, len(data.get("organic", [])) if isinstance(data, dict) else "")
os.makedirs("diag/output", exist_ok=True)
json.dump(out, open("diag/output/raw.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
