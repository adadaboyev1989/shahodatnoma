"""
e-shahodatnoma.uz -> "Arizalar" (/duplicates) jadvalini yig'ib olish dasturi.

Ish tartibi:
  1. Chrome oynasi ochiladi va sayt yuklanadi.
  2. Siz login/parolni o'zingiz kiritib, saytga kirasiz va Enter bosasiz.
  3. Dastur o'zi: Dublikat bo'limi -> 2026 tab -> Status = "Xulosa berildi".
  4. Filtrlangan jadvalning barcha sahifalarini varaqlab, arizalarni yig'adi
     va "natijalar" papkasiga Excel fayl qilib saqlaydi
     (Arizalar + Status/Tuman/Maktab/Muddat bo'yicha hisobot varaqlari).
  5. Har bir arizaga kirib, Xulosa blokidagi eng oxirgi "Xulosa berildi"
     bandining faylini "natijalar/xulosalar/<ariza raqami>.pdf" qilib yuklaydi.
     Avval yuklangan fayllar o'tkazib yuboriladi (dasturni qayta ishga
     tushirsangiz, to'xtagan joyidan davom etadi).

O'rnatish:
  pip install playwright pandas openpyxl
  python -m playwright install chrome   (agar kompyuterda Chrome bo'lmasa)

Ishga tushirish:
  python eshahodatnoma.py
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

SAYT = "https://e-shahodatnoma.uz/duplicates"
# exe holatida __file__ vaqtinchalik papkaga ishora qiladi, shuning uchun exe yonidagi papka olinadi
PAPKA = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
PROFIL = PAPKA / "chrome_profil"  # sessiya saqlanadi, har safar qayta kirish shart emas
NATIJALAR = PAPKA / "natijalar"
XULOSALAR = NATIJALAR / "xulosalar"

YIL = "2026"
STATUS = "Xulosa berildi"

GRID = ".rz-data-grid"
KEYINGI = ".rz-paginator-next"

# Brauzer ichida bajariladi: joriy sahifadagi qatorlarni oladi.
# Radzen har bir katakda ustun nomini (.rz-column-title) ham chiqaradi,
# shuning uchun ustunlar tartibiga emas, nomiga tayanamiz.
QATORLARNI_OLISH = """
grid => [...grid.querySelectorAll('tbody tr.rz-data-row')].map(tr => {
    const row = {};
    tr.querySelectorAll('td').forEach(td => {
        const nom = td.querySelector('.rz-column-title')?.textContent.trim();
        const qiymat = td.querySelector('.rz-cell-data');
        if (!nom || nom === '#' || !qiymat) return;
        row[nom] = qiymat.innerText.trim();
        const a = qiymat.querySelector('a[href]');
        if (a && nom === 'Ariza raqami') row['Havola'] = a.href;
    });
    return row;
})
"""


def jami_yozuv(page):
    matn = page.locator(".rz-paginator-summary").first.inner_text()
    raqam = "".join(ch for ch in matn if ch.isdigit())
    return int(raqam) if raqam else None


def sahifa_belgisi(page):
    """Joriy sahifadagi birinchi ariza raqami: sahifa almashganini aniqlash uchun."""
    birinchi = page.locator(f"{GRID} tbody tr.rz-data-row a[href]").first
    return birinchi.inner_text() if birinchi.count() else ""


def dublikat_filtrlash(page, yil=YIL, status=STATUS):
    """Dublikat bo'limi -> yil tabi -> Status ustuni filtri."""
    page.locator("a.rz-navigation-item-link", has_text="Dublikat").first.click()
    page.wait_for_url("**/duplicates*")
    page.locator(".rz-tabview-nav a[role=tab]", has_text=yil).first.click()
    page.wait_for_selector(f"{GRID} tbody tr.rz-data-row", timeout=60_000)
    oldingi_jami = jami_yozuv(page)

    # Filtr qatoridagi birinchi dropdown - Status ustuni (ikkinchisi - Operator)
    page.locator(f"{GRID} .filter-dropdown").first.click()
    page.locator(f'li.rz-dropdown-item[aria-label="{status}"]:visible').first.click()
    page.wait_for_function(
        """([sel, eski]) => {
            const el = document.querySelector(sel);
            return el && el.textContent.replace(/\\D/g, '') !== String(eski);
        }""",
        arg=[".rz-paginator-summary", oldingi_jami],
        timeout=60_000,
    )
    print(f"{yil} yil, Status = {status}: {jami_yozuv(page)} ta yozuv")


def barcha_arizalarni_yigish(page):
    page.wait_for_selector(f"{GRID} tbody tr.rz-data-row", timeout=60_000)
    jami = jami_yozuv(page)
    print(f"Jadvalda jami: {jami} ta yozuv")

    natija, korilgan = [], set()
    sahifa_no = 1
    while True:
        qatorlar = page.locator(GRID).first.evaluate(QATORLARNI_OLISH)
        for q in qatorlar:
            kalit = q.get("Ariza raqami")
            if kalit and kalit not in korilgan:
                korilgan.add(kalit)
                natija.append(q)
        print(f"\r{sahifa_no}-sahifa, yig'ildi: {len(natija)}/{jami}", end="", flush=True)

        keyingi = page.locator(KEYINGI).first
        if not keyingi.count() or "rz-state-disabled" in (keyingi.get_attribute("class") or ""):
            break
        oldingi = sahifa_belgisi(page)
        keyingi.click()
        page.wait_for_function(
            """([sel, eski]) => {
                const a = document.querySelector(sel);
                return a && a.textContent.trim() !== eski;
            }""",
            arg=[f"{GRID} tbody tr.rz-data-row a[href]", oldingi],
            timeout=60_000,
        )
        sahifa_no += 1
    print()
    return natija


def excelga_saqlash(arizalar):
    NATIJALAR.mkdir(exist_ok=True)
    fayl = NATIJALAR / f"arizalar_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    df = pd.DataFrame(arizalar)

    with pd.ExcelWriter(fayl) as writer:
        df.to_excel(writer, sheet_name="Arizalar", index=False)
        for ustun in ["Status", "Operator", "Tuman", "Maktab", "Muddat"]:
            if ustun in df:
                hisobot = (df[ustun].replace("", "(ko'rsatilmagan)")
                           .value_counts().rename_axis(ustun).reset_index(name="Soni"))
                hisobot.to_excel(writer, sheet_name=f"{ustun} bo'yicha"[:31], index=False)

    print(f"{len(df)} ta ariza saqlandi: {fayl}")
    if "Status" in df:
        print("\nStatus bo'yicha:")
        print(df["Status"].value_counts().to_string())


def xulosa_faylini_yuklash(page, ariza_raqami, havola):
    """Ariza sahifasidagi eng oxirgi "Xulosa berildi" bandining faylini yuklaydi."""
    mavjud = list(XULOSALAR.glob(f"{ariza_raqami}.*"))
    if mavjud:
        return "oldin yuklangan"

    page.goto(havola)
    page.wait_for_selector("ul.timeline li", timeout=60_000)
    bandlar = page.locator("ul.timeline li").filter(
        has=page.locator("span.text-success", has_text="Xulosa berildi")
    )
    if not bandlar.count():
        return "Xulosa berildi bandi topilmadi"

    # Bir nechta "Xulosa berildi" bo'lsa - vaqti eng so'nggisi ([YYYY-MM-DD HH:MM:SS]).
    # Vaqtlar teng bo'lsa yoki o'qilmasa, ro'yxatda pastrog'i olinadi.
    vaqtlar = bandlar.evaluate_all(
        "els => els.map(li => (li.querySelector('span.text-muted')?.textContent.match(/\\d{4}-\\d\\d-\\d\\d \\d\\d:\\d\\d:\\d\\d/) || [''])[0])"
    )
    eng_songgi = max(range(len(vaqtlar)), key=lambda i: (vaqtlar[i], i))
    fayl_havola = bandlar.nth(eng_songgi).locator("a").last
    if not fayl_havola.count():
        return "Xulosa faylsiz"

    asl_nom = fayl_havola.inner_text().strip()
    kengaytma = Path(asl_nom).suffix or ".pdf"
    manzil = XULOSALAR / f"{ariza_raqami}{kengaytma}"

    yangi_tablar = len(page.context.pages)
    try:
        with page.expect_download(timeout=30_000) as kutish:
            fayl_havola.click()
        kutish.value.save_as(manzil)
    except Exception:
        # Ba'zan fayl yuklanmasdan yangi tabda ochiladi: o'sha manzildan olamiz
        if len(page.context.pages) <= yangi_tablar:
            raise
        tab = page.context.pages[-1]
        tab.wait_for_load_state()
        manzil.write_bytes(page.context.request.get(tab.url).body())
        tab.close()
    return f"yuklandi ({asl_nom})"


def xulosalarni_yuklash(page, arizalar):
    XULOSALAR.mkdir(parents=True, exist_ok=True)
    hisobot = []
    for i, ariza in enumerate(arizalar, 1):
        raqam, havola = ariza.get("Ariza raqami"), ariza.get("Havola")
        try:
            holat = xulosa_faylini_yuklash(page, raqam, havola)
        except Exception as xato:
            holat = f"xatolik: {str(xato).splitlines()[0]}"
        print(f"[{i}/{len(arizalar)}] {raqam}: {holat}")
        hisobot.append({"Ariza raqami": raqam, "Holat": holat})

    fayl = XULOSALAR / f"hisobot_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    pd.DataFrame(hisobot).to_excel(fayl, index=False)
    print(f"\nXulosalar papkasi: {XULOSALAR}\nYuklash hisoboti: {fayl}")


def brauzerni_ochish(p):
    """Kompyuterdagi Chrome ochiladi; u bo'lmasa - Windows bilan keladigan Edge."""
    for kanal in ("chrome", "msedge"):
        try:
            return p.chromium.launch_persistent_context(
                str(PROFIL), channel=kanal, headless=False, no_viewport=True,
                args=["--start-maximized"],
            )
        except Exception as xato:
            oxirgi_xato = xato
    raise RuntimeError(f"Chrome ham, Edge ham topilmadi: {oxirgi_xato}")


def main():
    with sync_playwright() as p:
        context = brauzerni_ochish(p)
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(SAYT)

        print("Brauzer ochildi. Iltimos, login va parolingizni kiritib saytga kiring.")
        input("Kirib bo'lgach, Enter bosing... ")
        page = context.pages[-1]  # oxirgi ochilgan tab
        try:
            dublikat_filtrlash(page)
            arizalar = barcha_arizalarni_yigish(page)
            excelga_saqlash(arizalar)
            xulosalarni_yuklash(page, arizalar)
        except Exception as xato:
            print(f"\nXatolik: {xato}")
        input("\nYopish uchun Enter bosing... ")

        context.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as xato:
        # exe oynasi xatoni ko'rsatmasdan yopilib qolmasligi uchun
        print(f"\nXatolik: {xato}")
        input("Yopish uchun Enter bosing... ")
