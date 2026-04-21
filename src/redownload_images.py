"""Re-download product images by variant + largest resolution.

Groups Amazon images by variant (MAIN, PT01-PT08) and picks the highest-resolution
image per variant. Outputs to data/images_v2/{sku}/ for clean URLs.
"""
import json, os, urllib.request
from collections import defaultdict

os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

OUT_DIR = "data/images_v2"
os.makedirs(OUT_DIR, exist_ok=True)

with open("data/amz_catalog.json", "r", encoding="utf-8") as f:
    catalog = json.load(f)

manifest = []
for entry in catalog:
    sku = entry["sku"]
    if "error" in entry:
        print(f"  skip {sku} (error)")
        continue
    data = entry["data"]
    imgs = data.get("images", [])
    # Flatten, filter by size, group by variant
    per_variant = defaultdict(list)
    for group in imgs:
        for img in group.get("images", []):
            variant = img.get("variant", "MAIN")
            w, h = img.get("width", 0), img.get("height", 0)
            link = img.get("link", "")
            if not link:
                continue
            per_variant[variant].append((w * h, w, h, link))

    # For each variant, pick the largest image
    # Order: MAIN, PT01, PT02, ..., PT08
    def variant_sort_key(v):
        if v == "MAIN":
            return 0
        if v.startswith("PT"):
            try:
                return int(v[2:])
            except:
                return 999
        return 1000

    chosen = []
    for variant in sorted(per_variant.keys(), key=variant_sort_key):
        candidates = per_variant[variant]
        # pick largest
        candidates.sort(reverse=True)
        _, w, h, link = candidates[0]
        chosen.append((variant, w, h, link))

    # Take up to 8 (1 main + 7 secondary)
    chosen = chosen[:8]

    sku_dir = f"{OUT_DIR}/{sku}"
    os.makedirs(sku_dir, exist_ok=True)
    local_paths = []
    for i, (variant, w, h, link) in enumerate(chosen):
        fname = f"{sku}_{i+1}.jpg"
        path = f"{sku_dir}/{fname}"
        if not os.path.exists(path) or os.path.getsize(path) < 10000:
            try:
                req = urllib.request.Request(link, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=30) as resp:
                    data_bytes = resp.read()
                with open(path, "wb") as out:
                    out.write(data_bytes)
            except Exception as e:
                print(f"  FAIL {sku} {variant}: {e}")
                continue
        local_paths.append(path)

    print(f"{sku:<16}: {len(chosen)} images ({', '.join(v for v,_,_,_ in chosen)})")
    manifest.append({"sku": sku, "variants": [{"variant": v, "w": w, "h": h, "url": l} for v,w,h,l in chosen]})

with open("data/image_manifest_v2.json", "w") as f:
    json.dump(manifest, f, indent=2)
print(f"\nSaved manifest for {len(manifest)} products")
