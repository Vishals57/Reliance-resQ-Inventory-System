import qrcode
import os

def generate_article_qr(art_no, name):
    """
    Backward compatible wrapper.
    Project now uses ONLY one QR type: Box QR (full details).
    """
    return generate_box_qr(art_no=art_no, part_name=name, tax_invoice_no="NA", sp="-", qty=1, sr_list=[1])

def generate_unit_qr(art_no: str, sr_no: int, part_name: str, tax_invoice_no: str, sp):
    """
    Backward compatible wrapper.
    Project now uses ONLY one QR type: Box QR (full details).
    """
    return generate_box_qr(art_no=art_no, part_name=part_name, tax_invoice_no=tax_invoice_no, sp=sp, qty=1, sr_list=[sr_no])

def generate_box_qr(art_no: str, part_name: str, tax_invoice_no: str, sp, qty: int, sr_list=None):
    """
    Single box QR for the whole inward shipment.
    Mobile scan shows full details; app can still extract Art Code.
    """
    art_no = str(art_no).strip()
    part_name = "" if part_name is None else str(part_name).strip()
    tax_invoice_no = "" if tax_invoice_no is None else str(tax_invoice_no).strip()
    try:
        qty = int(qty)
    except Exception:
        qty = 1
    if qty < 1:
        qty = 1

    sp_txt = "-" if sp is None else str(sp)

    sr_text = "-"
    if sr_list:
        try:
            srs = [int(x) for x in sr_list]
            if srs:
                srs = sorted(set(srs))
                sr_text = f"{srs[0]}..{srs[-1]}" if len(srs) > 1 else str(srs[0])
        except Exception:
            sr_text = "-"

    payload = "\n".join([
        f"ART_CODE: {art_no}",
        f"PART_NAME: {part_name}",
        f"INVOICE_NO: {tax_invoice_no or '-'}",
        f"SELLING_PRICE: {sp_txt}",
        f"QTY: {qty}",
        f"SR_NOS: {sr_text}",
    ])

    if not os.path.exists("Box_QRs"):
        os.makedirs("Box_QRs")

    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(payload)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    safe_inv = "".join(ch for ch in (tax_invoice_no or "NA") if ch.isalnum() or ch in ("-", "_"))[:30]
    file_path = f"Box_QRs/QR_{art_no}_{safe_inv}.png"
    img.save(file_path)
    return file_path