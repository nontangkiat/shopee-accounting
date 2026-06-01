from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import anthropic, base64, json, os, tempfile, uuid
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
import io

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

PROMPT = """Extract Shopee purchase data from this screenshot. Return ONLY raw JSON, no markdown.

{"orders":[{"order_id":"","date":"","shop_name":"","item_name":"","quantity":1,"unit_price":0,"total_price":0,"net_total":0,"shipping_fee":0,"discount":0,"payment_method":"","status":""}]}

- Extract every item visible
- Prices as numbers in THB (no symbols)  
- Missing fields: "" or 0
- net_total = actual amount paid"""

def extract_data(image_bytes: bytes, media_type: str) -> list:
    b64 = base64.standard_b64encode(image_bytes).decode()
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
            {"type": "text", "text": PROMPT}
        ]}]
    )
    text = msg.content[0].text
    s, e = text.index("{"), text.rindex("}")
    return json.loads(text[s:e+1]).get("orders", [])

def make_excel(all_orders: list, image_map: dict) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Shopee Orders"

    # Styles
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hdr_fill = PatternFill("solid", start_color="EE4D2D")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill = PatternFill("solid", start_color="FFF8F7")

    headers = [("หลักฐาน",18),("Order ID",22),("วันที่",13),("ร้านค้า",26),
               ("ชื่อสินค้า",34),("จำนวน",8),("ราคา/หน่วย",12),("รวม",11),
               ("ยอดสุทธิ",11),("ค่าส่ง",9),("สถานะ",20)]

    for c, (h, w) in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = hdr_align; cell.border = bdr
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 28

    ROW_H = 80
    for ri, order in enumerate(all_orders, 2):
        ws.row_dimensions[ri].height = ROW_H
        row_fill = alt_fill if ri % 2 == 0 else None
        vals = ["", order.get("order_id",""), order.get("date",""),
                order.get("shop_name",""), order.get("item_name",""),
                order.get("quantity",0), order.get("unit_price",0),
                order.get("total_price",0), order.get("net_total",0),
                order.get("shipping_fee",0), order.get("status","")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(ri, c, v)
            cell.font = Font(name="Arial", size=10)
            cell.border = bdr
            if row_fill: cell.fill = row_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                horizontal="center" if c in (1,6) else "right" if c in (7,8,9,10) else "left")

        # Embed full-size image (fit to row)
        img_bytes = image_map.get(order.get("_file",""))
        if img_bytes:
            with PILImage.open(io.BytesIO(img_bytes)) as im:
                im = im.convert("RGB")
                # Fit height to row (ROW_H pts ≈ ROW_H*1.33 px), keep aspect
                target_h = int(ROW_H * 1.2)
                ratio = target_h / im.height
                target_w = int(im.width * ratio)
                im = im.resize((target_w, target_h), PILImage.LANCZOS)
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=90)
                buf.seek(0)
            tmp = f"/tmp/img_{ri}.jpg"
            with open(tmp, "wb") as f: f.write(buf.read())
            xl_img = XLImage(tmp)
            xl_img.width = target_w
            xl_img.height = target_h
            ws.add_image(xl_img, f"A{ri}")

    # Summary
    sr = len(all_orders) + 2
    ws.cell(sr, 4, "รวมทั้งหมด").font = Font(name="Arial", bold=True, size=11)
    ws.cell(sr, 8, f"=SUM(H2:H{sr-1})").font = Font(name="Arial", bold=True)
    ws.cell(sr, 9, f"=SUM(I2:I{sr-1})").font = Font(name="Arial", bold=True)
    sum_fill = PatternFill("solid", start_color="FFE5E0")
    for c in range(1, 12):
        ws.cell(sr, c).border = bdr
        ws.cell(sr, c).fill = sum_fill
    ws.freeze_panes = "B2"

    out = f"/tmp/shopee_{uuid.uuid4().hex[:8]}.xlsx"
    wb.save(out)
    return out

@app.post("/process")
async def process(files: list[UploadFile] = File(...)):
    all_orders = []
    image_map = {}
    for f in files:
        data = await f.read()
        mt = f.content_type or "image/jpeg"
        if not mt.startswith("image/"):
            mt = "image/jpeg"
        orders = extract_data(data, mt)
        for o in orders:
            o["_file"] = f.filename
        image_map[f.filename] = data
        all_orders.extend(orders)
    path = make_excel(all_orders, image_map)
    return FileResponse(path, filename="shopee_orders.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/health")
def health(): return {"status": "ok"}
