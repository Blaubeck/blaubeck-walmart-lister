"""Fill Walmart template by driving Excel directly via COM.

This preserves ALL of Walmart's built-in validations, dropdowns, formulas,
and formatting that openpyxl would otherwise damage.
"""
import json, csv, os, shutil, time
import win32com.client as win32
from pathlib import Path

os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

TEMPLATE = os.path.abspath("data/walmart_template.xlsx")
OUTPUT = os.path.abspath("data/walmart_filled.xlsx")
IMG_BASE = "https://raw.githubusercontent.com/Blaubeck/blaubeck-walmart-lister/master/data/images_v2"

PRODUCT_TYPE = {
    "MagHolder02": "Electronics Stands",
    "MagHolder04": "Electronics Stands",
    "MagHolder06": "Electronics Stands",
    "PadelHolder01": "Electronics Stands",  # not Car Mounts - padel not automotive
    "StickerPack01": "Stickers",
    "StickerPack02": "Stickers",
    "StickerPack03": "Stickers",
    "StickerPack04": "Stickers",
    "StickerPack05": "Stickers",
    "Adhesive01": "Hardware Hooks",
    "PB-R5M6-S3DY": "Sheet Protectors",
}

FULFILLMENT_CENTER_ID = "10001071862"  # Our seller-fulfilled warehouse

SKU_EXTRAS = {
    "MagHolder02": {"color": "Black", "material": "Aluminum", "attachmentStyle": "Clip-On"},
    "MagHolder04": {"color": "Black", "material": "Zinc Alloy", "attachmentStyle": "Magnetic"},
    "MagHolder06": {"color": "Black", "material": "Carbon Fiber", "attachmentStyle": "Magnetic"},
    "PadelHolder01": {"color": "Black", "material": "Aluminum", "attachmentStyle": "Clip-On"},
    "StickerPack01": {"color": "Multi", "material": "Vinyl", "sticker_type": "Vinyl Stickers", "pieceCount": 34, "pattern": "Graphic"},
    "StickerPack02": {"color": "Neon", "material": "Vinyl", "sticker_type": "Vinyl Stickers", "pieceCount": 34, "pattern": "Graphic"},
    "StickerPack03": {"color": "Multi", "material": "Vinyl", "sticker_type": "Vinyl Stickers", "pieceCount": 45, "pattern": "Graphic"},
    "StickerPack04": {"color": "Multi", "material": "Vinyl", "sticker_type": "Vinyl Stickers", "pieceCount": 43, "pattern": "Graphic"},
    "StickerPack05": {"color": "Multi", "material": "Vinyl", "sticker_type": "Vinyl Stickers", "pieceCount": 46, "pattern": "Graphic"},
    "Adhesive01": {"color": "White", "material": "Plastic", "hardware_hook_type": "Adhesive Hooks", "installationType": "Adhesive", "pieceCount": 4},
    "PB-R5M6-S3DY": {"color": "Clear", "material": "Vinyl", "numberOfSheets": 20, "pieceCount": 20},
}

# 1-indexed column positions on "Product Content And Site Exp" sheet
C = {
    "sku": 4, "specProductType": 5, "productIdType": 6, "productId": 7,
    "productName": 8, "brand": 9, "price": 10, "ShippingWeight": 11,
    "country_of_origin_substantial_transformation": 12, "quantity": 14,
    "shortDescription": 16,
    "keyFeatures1": 17, "keyFeatures2": 18, "keyFeatures3": 19, "keyFeatures4": 20,
    "mainImageUrl": 21,
    "isProp65WarningRequired": 22, "condition": 23, "has_written_warranty": 24,
    "productNetContentUnit": 25, "productNetContentMeasure": 26,
    "assembledProductHeight_measure": 27, "assembledProductHeight_unit": 28,
    "assembledProductWidth_measure": 29, "assembledProductWidth_unit": 30,
    "color": 31, "has_nrtl_listing_certification": 32,
    "smallPartsWarnings": 33,
    "productSecondaryImageURL1": 36, "productSecondaryImageURL2": 37, "productSecondaryImageURL3": 38,
    "assembledProductLength_measure": 40, "assembledProductLength_unit": 41,
    "assembledProductWeight_measure": 42, "assembledProductWeight_unit": 43,
    "attachmentStyle": 44,
    "hardware_hook_type": 63, "installationType": 64,
    "fulfillmentCenterID": 13,
    "manufacturer": 67, "manufacturerPartNumber": 68,
    "material": 69, "modelNumber": 83, "mountType": 84,
    "pieceCount": 90, "numberOfSheets": 91,
    "pattern": 98, "sticker_type": 114,
}


def image_urls(sku):
    d = f"data/images_v2/{sku}"
    if not os.path.isdir(d):
        return []
    # Sort by the numeric suffix (SKU_1.jpg, SKU_2.jpg ... SKU_10.jpg)
    def key(fname):
        stem = fname.rsplit(".", 1)[0]
        try:
            return int(stem.rsplit("_", 1)[-1])
        except:
            return 999
    files = sorted(os.listdir(d), key=key)
    return [f"{IMG_BASE}/{sku}/{f}" for f in files]


def load_data():
    with open("data/review_master.csv", "r", encoding="utf-8") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    included = {r["AMZ_SKU"]: r for r in rows if r["INCLUDE"].upper() == "YES"}
    with open("data/upcs.json", "r", encoding="utf-8") as f:
        upcs_raw = json.load(f)
    upcs = {k: v for k, v in upcs_raw.items() if not k.startswith("_")}
    with open("data/walmart_copy.json", "r", encoding="utf-8") as f:
        copy = json.load(f)
    return included, upcs, copy


def build_records(included, upcs, copy):
    records = []
    for sku, row in included.items():
        if sku not in copy: continue
        c = copy[sku]
        extras = SKU_EXTRAS.get(sku, {})
        imgs = image_urls(sku)
        main_img = imgs[0] if imgs else ""
        secondary = imgs[1:]
        product_type = PRODUCT_TYPE.get(sku, "Electronics Stands")
        upc = upcs.get(sku, "")
        gtin14 = ("0" + upc) if len(upc) == 13 else upc
        kf = list(c["keyFeatures"])
        while len(kf) < 4: kf.append("")

        rec = {
            "sku": sku,
            "specProductType": product_type,
            "productIdType": "GTIN",
            "productId": gtin14,
            "productName": c["title"],
            "brand": "BLAUBECK",
            "price": float(row["Price_USD"]),
            "ShippingWeight": 1.0,
            "country_of_origin_substantial_transformation": "China",
            "quantity": int(row["Stock"]) if int(row["Stock"]) > 0 else 100,
            "shortDescription": c["shortDescription"],
            "keyFeatures1": kf[0], "keyFeatures2": kf[1], "keyFeatures3": kf[2], "keyFeatures4": kf[3],
            "mainImageUrl": main_img,
            "condition": "New",
            "isProp65WarningRequired": "No",
            "has_written_warranty": "No",
            "has_nrtl_listing_certification": "No",
            "productNetContentMeasure": extras.get("pieceCount", 1),
            "productNetContentUnit": "Count",
            "fulfillmentCenterID": FULFILLMENT_CENTER_ID,
            "manufacturer": "BLAUBECK",
            "manufacturerPartNumber": sku,
            "modelNumber": sku,
            "assembledProductHeight_measure": 4,
            "assembledProductHeight_unit": "in",
            "assembledProductWidth_measure": 4,
            "assembledProductWidth_unit": "in",
            "assembledProductLength_measure": 4,
            "assembledProductLength_unit": "in",
            "assembledProductWeight_measure": 0.5,
            "assembledProductWeight_unit": "lb",
        }
        if len(secondary) >= 1: rec["productSecondaryImageURL1"] = secondary[0]
        if len(secondary) >= 2: rec["productSecondaryImageURL2"] = secondary[1]
        if len(secondary) >= 3: rec["productSecondaryImageURL3"] = secondary[2]
        if "color" in extras: rec["color"] = extras["color"]
        if "material" in extras: rec["material"] = extras["material"]
        if product_type in ("Electronics Stands", "Car Mounts"):
            if "mountType" in extras: rec["mountType"] = extras["mountType"]
            if "attachmentStyle" in extras: rec["attachmentStyle"] = extras["attachmentStyle"]
        if product_type == "Stickers":
            rec["sticker_type"] = extras.get("sticker_type", "Decorative")
            rec["pieceCount"] = extras.get("pieceCount", 30)
            rec["pattern"] = extras.get("pattern", "Graphic")
            rec["smallPartsWarnings"] = "0 - No warning applicable"
        if product_type == "Hardware Hooks":
            rec["hardware_hook_type"] = extras.get("hardware_hook_type", "Adhesive")
            rec["installationType"] = extras.get("installationType", "Adhesive")
            rec["pieceCount"] = extras.get("pieceCount", 4)
            rec["smallPartsWarnings"] = "0 - No warning applicable"
        if product_type == "Sheet Protectors":
            rec["numberOfSheets"] = extras.get("numberOfSheets", 20)
            rec["pieceCount"] = extras.get("pieceCount", 20)
        records.append(rec)
    return records


def main():
    # Start with a fresh copy of the template
    shutil.copy(TEMPLATE, OUTPUT)

    included, upcs, copy = load_data()
    records = build_records(included, upcs, copy)
    print(f"Built {len(records)} records")

    # Drive Excel via COM
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(OUTPUT)
        ws = wb.Worksheets("Product Content And Site Exp")

        start_row = 8
        for i, rec in enumerate(records):
            row_num = start_row + i
            for field, value in rec.items():
                if field not in C:
                    continue
                col = C[field]
                ws.Cells(row_num, col).Value = value
            print(f"  row {row_num}: {rec['sku']:<16} -> {rec['specProductType']}")

        wb.Save()
        wb.Close(False)
    finally:
        excel.Quit()

    print(f"\nSaved {OUTPUT}")


if __name__ == "__main__":
    main()
