"""Build the Walmart MP_ITEM feed from review_master.csv + walmart_copy.json + images.

NOTE: Walmart's MP_ITEM JSON schema is product-type specific. This generator outputs a
generic-structure feed you can validate against Walmart's Get Spec API on Monday before
submission. Product type mapping is conservative (general-purpose categories).

Run: python src/build_feed.py
Output: data/walmart_feed.json
"""
import csv, json, os
from walmart_copy import COPY

# Load UPCs (populated from GS1 UK on Monday)
with open("data/upcs.json", "r", encoding="utf-8") as f:
    _upcs_raw = json.load(f)
UPCS = {k: v for k, v in _upcs_raw.items() if not k.startswith("_")}

IMG_BASE = "https://raw.githubusercontent.com/Blaubeck/blaubeck-walmart-lister/master/data/images"

# Product type mapping - refine Monday after checking Walmart taxonomy
PRODUCT_TYPE = {
    "MagHolder02": "Cell Phone Mounts & Stands",
    "MagHolder04": "Cell Phone Mounts & Stands",
    "MagHolder06": "Cell Phone Mounts & Stands",
    "PadelHolder01": "Cell Phone Mounts & Stands",
    "StickerPack01": "Stickers & Decals",
    "StickerPack02": "Stickers & Decals",
    "StickerPack03": "Stickers & Decals",
    "StickerPack04": "Stickers & Decals",
    "StickerPack05": "Stickers & Decals",
    "Adhesive01": "Command Hooks",
    "PB-R5M6-S3DY": "Home Decor Accents",
}
DEFAULT_PRODUCT_TYPE = "Home Decor Accents"  # for MedalChalk01 etc


def image_urls(sku):
    d = f"data/images/{sku}"
    if not os.path.isdir(d):
        return []
    files = sorted(os.listdir(d))
    return [f"{IMG_BASE}/{sku}/{f}" for f in files]


def build_item(row, copy):
    sku = row["AMZ_SKU"]
    wm_sku = row["WALMART_SKU_OVERRIDE"] or sku
    imgs = image_urls(sku)
    main_img = imgs[0] if imgs else ""
    secondary = imgs[1:]

    upc = UPCS.get(sku, "UPC_PLACEHOLDER_FILL_MONDAY")
    if "PLACEHOLDER" in upc:
        upc = "UPC_PLACEHOLDER_FILL_MONDAY"

    return {
        "sku": wm_sku,
        "productIdentifiers": {
            "productIdType": "UPC",
            "productId": upc,
        },
        "productName": copy["title"],
        "brand": "BLAUBECK",
        "productCategory": PRODUCT_TYPE.get(sku, DEFAULT_PRODUCT_TYPE),
        "shortDescription": copy["shortDescription"],
        "keyFeatures": copy["keyFeatures"],
        "description": copy["description"],
        "keywords": copy["keywords"],
        "mainImageUrl": main_img,
        "productSecondaryImageURL": secondary,
        "price": {
            "currency": "USD",
            "amount": float(row["Price_USD"]),
        },
        "shippingWeight": 1.0,  # placeholder lb - update per product
        "shippingWeightUnit": "lb",
        "manufacturer": "BLAUBECK",
        "manufacturerPartNumber": wm_sku,
        "countryOfOriginAssembly": "CN",  # most of our products - update if different per SKU
    }


def main():
    with open("data/review_master.csv", "r", encoding="utf-8") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    included = [r for r in rows if r["INCLUDE"].upper() == "YES"]

    items = []
    for r in included:
        sku = r["AMZ_SKU"]
        if sku not in COPY:
            print(f"  SKIP {sku}: no copy defined")
            continue
        items.append(build_item(r, COPY[sku]))

    feed = {
        "MPItemFeedHeader": {
            "sellingChannel": "marketplace",
            "locale": "en_US",
            "version": "5.0",
        },
        "MPItem": items,
    }

    with open("data/walmart_feed.json", "w", encoding="utf-8") as f:
        json.dump(feed, f, indent=2, ensure_ascii=False)

    missing_upc = [i for i in items if "PLACEHOLDER" in i["productIdentifiers"]["productId"]]
    print(f"Built feed with {len(items)} items")
    print(f"Items needing UPC Monday: {len(missing_upc)}")
    print("Output: data/walmart_feed.json")
    print()
    for item in items:
        upc_status = "PLACEHOLDER" if "PLACEHOLDER" in item["productIdentifiers"]["productId"] else item["productIdentifiers"]["productId"]
        print(f"  {item['sku']:<16} UPC={upc_status:<20} {len(item['productSecondaryImageURL'])+1} imgs  ${item['price']['amount']}")


if __name__ == "__main__":
    main()
