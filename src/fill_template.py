"""Fill Walmart's XLSX template with all required + product-type-specific fields.

Read:  data/walmart_template.xlsx (downloaded from Seller Center)
Write: data/walmart_filled.xlsx (ready for Seller Center upload)
"""
import json, csv, os, shutil, openpyxl

os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

TEMPLATE = "data/walmart_template.xlsx"
OUTPUT = "data/walmart_filled.xlsx"
IMG_BASE = "https://raw.githubusercontent.com/Blaubeck/blaubeck-walmart-lister/master/data/images"

# Product type mapping
PRODUCT_TYPE = {
    "MagHolder02": "Electronics Stands",
    "MagHolder04": "Electronics Stands",
    "MagHolder06": "Electronics Stands",
    "PadelHolder01": "Car Mounts",
    "StickerPack01": "Stickers",
    "StickerPack02": "Stickers",
    "StickerPack03": "Stickers",
    "StickerPack04": "Stickers",
    "StickerPack05": "Stickers",
    "Adhesive01": "Hardware Hooks",
    "PB-R5M6-S3DY": "Sheet Protectors",
}

# Per-SKU extras (color, piece count, sticker type, etc.)
SKU_EXTRAS = {
    "MagHolder02": {"color": "Black", "material": "Aluminum", "mountType": "Clamp", "attachmentStyle": "Clamp"},
    "MagHolder04": {"color": "Black", "material": "Zinc Alloy", "mountType": "Magnetic", "attachmentStyle": "Magnetic"},
    "MagHolder06": {"color": "Black", "material": "Carbon Fiber", "mountType": "Magnetic", "attachmentStyle": "Magnetic"},
    "PadelHolder01": {"color": "Black", "material": "Aluminum", "mountType": "Suction Cup", "attachmentStyle": "Suction Cup"},
    "StickerPack01": {"color": "Multi", "material": "Vinyl", "sticker_type": "Decorative", "pieceCount": 34, "pattern": "Graphic"},
    "StickerPack02": {"color": "Neon", "material": "Vinyl", "sticker_type": "Decorative", "pieceCount": 34, "pattern": "Graphic"},
    "StickerPack03": {"color": "Multi", "material": "Vinyl", "sticker_type": "Decorative", "pieceCount": 45, "pattern": "Graphic"},
    "StickerPack04": {"color": "Multi", "material": "Vinyl", "sticker_type": "Decorative", "pieceCount": 43, "pattern": "Graphic"},
    "StickerPack05": {"color": "Multi", "material": "Vinyl", "sticker_type": "Decorative", "pieceCount": 46, "pattern": "Graphic"},
    "Adhesive01": {"color": "White", "material": "Plastic", "hardware_hook_type": "Adhesive", "installationType": "Adhesive", "pieceCount": 4},
    "PB-R5M6-S3DY": {"color": "Clear", "material": "Vinyl", "numberOfSheets": 20, "pieceCount": 20},
}

# Column positions (1-indexed, from inspection)
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
    "smallPartsWarnings": 33, "accessoriesIncluded": 34, "features": 35,
    "productSecondaryImageURL1": 36, "productSecondaryImageURL2": 37, "productSecondaryImageURL3": 38,
    "assembledProductLength_measure": 40, "assembledProductLength_unit": 41,
    "assembledProductWeight_measure": 42, "assembledProductWeight_unit": 43,
    "attachmentStyle": 44,
    "hardware_hook_type": 63, "installationType": 64,
    "items_included": 65,
    "manufacturer": 67, "manufacturerPartNumber": 68,
    "material": 69,
    "modelNumber": 83, "mountType": 84,
    "pieceCount": 90, "numberOfSheets": 91,
    "sticker_type": 114,
    "pattern": 98,
    "shape": 109, "size": 110,
    "count": 119,  # Total Count for packs
}


def image_urls(sku):
    d = f"data/images/{sku}"
    if not os.path.isdir(d):
        return []
    return sorted([f"{IMG_BASE}/{sku}/{f}" for f in os.listdir(d)])


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


def set_cell(ws, row, col_key, value):
    col = C[col_key]
    ws.cell(row=row, column=col, value=value)


def main():
    shutil.copy(TEMPLATE, OUTPUT)
    wb = openpyxl.load_workbook(OUTPUT)
    ws = wb["Product Content And Site Exp"]

    included, upcs, copy = load_data()

    start_row = 8
    i = 0
    for sku, row in included.items():
        if sku not in copy: continue
        c = copy[sku]
        extras = SKU_EXTRAS.get(sku, {})
        imgs = image_urls(sku)
        main_img = imgs[0] if imgs else ""
        secondary = imgs[1:]

        product_type = PRODUCT_TYPE.get(sku, "Electronics Stands")
        upc = upcs.get(sku, "")

        kf = list(c["keyFeatures"])
        while len(kf) < 4: kf.append("")

        r = start_row + i

        # Core required
        set_cell(ws, r, "sku", sku)
        set_cell(ws, r, "specProductType", product_type)
        set_cell(ws, r, "productIdType", "UPC")
        set_cell(ws, r, "productId", upc)
        set_cell(ws, r, "productName", c["title"])
        set_cell(ws, r, "brand", "BLAUBECK")
        set_cell(ws, r, "price", float(row["Price_USD"]))
        set_cell(ws, r, "ShippingWeight", 1.0)
        set_cell(ws, r, "country_of_origin_substantial_transformation", "China")
        set_cell(ws, r, "quantity", int(row["Stock"]) if int(row["Stock"]) > 0 else 100)
        set_cell(ws, r, "shortDescription", c["shortDescription"])
        set_cell(ws, r, "keyFeatures1", kf[0])
        set_cell(ws, r, "keyFeatures2", kf[1])
        set_cell(ws, r, "keyFeatures3", kf[2])
        set_cell(ws, r, "keyFeatures4", kf[3])
        set_cell(ws, r, "mainImageUrl", main_img)
        set_cell(ws, r, "condition", "New")
        set_cell(ws, r, "isProp65WarningRequired", "No")
        set_cell(ws, r, "has_written_warranty", "No")
        set_cell(ws, r, "has_nrtl_listing_certification", "No")

        # Net content (1 per pack - except stickers & hooks have counts)
        piece_count = extras.get("pieceCount", 1)
        set_cell(ws, r, "productNetContentMeasure", piece_count)
        set_cell(ws, r, "productNetContentUnit", "Count")

        # Secondary images
        for k, url in enumerate(secondary[:3]):
            set_cell(ws, r, f"productSecondaryImageURL{k+1}", url)

        # Manufacturer
        set_cell(ws, r, "manufacturer", "BLAUBECK")
        set_cell(ws, r, "manufacturerPartNumber", sku)
        set_cell(ws, r, "modelNumber", sku)

        # Common descriptive
        if "color" in extras:
            set_cell(ws, r, "color", extras["color"])
        if "material" in extras:
            set_cell(ws, r, "material", extras["material"])

        # Assembled dimensions (small defaults - user can refine later)
        set_cell(ws, r, "assembledProductHeight_measure", 4)
        set_cell(ws, r, "assembledProductHeight_unit", "inch")
        set_cell(ws, r, "assembledProductWidth_measure", 4)
        set_cell(ws, r, "assembledProductWidth_unit", "inch")
        set_cell(ws, r, "assembledProductLength_measure", 4)
        set_cell(ws, r, "assembledProductLength_unit", "inch")
        set_cell(ws, r, "assembledProductWeight_measure", 0.5)
        set_cell(ws, r, "assembledProductWeight_unit", "lb")

        # PT-specific
        if product_type in ("Electronics Stands", "Car Mounts"):
            if "mountType" in extras:
                set_cell(ws, r, "mountType", extras["mountType"])
            if "attachmentStyle" in extras:
                set_cell(ws, r, "attachmentStyle", extras["attachmentStyle"])

        if product_type == "Stickers":
            set_cell(ws, r, "sticker_type", extras.get("sticker_type", "Decorative"))
            set_cell(ws, r, "pieceCount", extras.get("pieceCount", 30))
            set_cell(ws, r, "pattern", extras.get("pattern", "Graphic"))
            set_cell(ws, r, "smallPartsWarnings", "No Warning Applicable")

        if product_type == "Hardware Hooks":
            set_cell(ws, r, "hardware_hook_type", extras.get("hardware_hook_type", "Adhesive"))
            set_cell(ws, r, "installationType", extras.get("installationType", "Adhesive"))
            set_cell(ws, r, "pieceCount", extras.get("pieceCount", 4))
            set_cell(ws, r, "smallPartsWarnings", "No Warning Applicable")

        if product_type == "Sheet Protectors":
            set_cell(ws, r, "numberOfSheets", extras.get("numberOfSheets", 20))
            set_cell(ws, r, "pieceCount", extras.get("pieceCount", 20))

        print(f"  row {r}: {sku:<16} -> {product_type} ({piece_count}pc, color={extras.get('color','-')})")
        i += 1

    wb.save(OUTPUT)
    print(f"\nSaved {OUTPUT} with {i} rows")


if __name__ == "__main__":
    main()
