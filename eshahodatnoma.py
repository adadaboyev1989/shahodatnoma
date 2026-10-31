"""
e-shahodatnoma.uz -> "Arizalar" (/duplicates) jadvalini yig'ib olish dasturi.

Ish tartibi:
  1. Chrome oynasi ochiladi va sayt yuklanadi.
  2. Siz login/parolni o'zingiz kiritib, saytga kirasiz va Enter bosasiz.
  3. Dastur o'zi: Dublikat bo'limi -> 2026 tab -> Status = "Xulosa berildi".
  4. Filtrlangan jadvalning barcha sahifalarini varaqlab, arizalarni yig'adi
     va "natijalar" papkasiga Excel fayl qilib saqlaydi
     (Arizalar + Status/Tuman/Maktab/Muddat bo'yicha hisobot varaqlari).

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
PROFIL = Path(__file__).with_name("chrome_profil")  # sessiya saqlanadi, har safar qayta kirish shart emas
NATIJALAR = Path(__file__).with_name("natijalar")

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


def main():
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            str(PROFIL), channel="chrome", headless=False, no_viewport=True,
            args=["--start-maximized"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(SAYT)

        print("Chrome ochildi. Iltimos, login va parolingizni kiritib saytga kiring.")
        input("Kirib bo'lgach, Enter bosing... ")
        page = context.pages[-1]  # oxirgi ochilgan tab
        try:
            dublikat_filtrlash(page)
            excelga_saqlash(barcha_arizalarni_yigish(page))
        except Exception as xato:
            print(f"\nXatolik: {xato}")
        input("\nYopish uchun Enter bosing... ")

        context.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
