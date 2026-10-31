"""
e-shahodatnoma.uz -> "Arizalar" (/duplicates) jadvalini yig'ib olish dasturi.

Ish tartibi:
  1. Chrome oynasi ochiladi va sayt yuklanadi.
  2. Siz login/parolni o'zingiz kiritib, saytga kirasiz. Dastur kirganingizni
     o'zi sezadi va davom etadi (Enter bosish shart emas).
  3. Dastur o'zi: Dublikat bo'limi -> 2026 tab -> Status = "Xulosa berildi".
  4. Filtrlangan jadvalning barcha sahifalarini varaqlab, arizalarni yig'adi
     va "natijalar" papkasiga Excel fayl qilib saqlaydi
     (Arizalar + Yil/Status/Tuman/Maktab/Muddat bo'yicha hisobot varaqlari).
  5. Har bir arizaga kirib, Xulosa blokidagi eng oxirgi "Xulosa berildi"
     bandining faylini "Sana / vaqt" ustunidagi yil bo'yicha papkaga yuklaydi
     (2026 tabida 2025-yil sanali arizalar ham bor - ular alohida papkaga tushadi):
     "natijalar/xulosalar/<yil>/<ariza raqami>.pdf".
     Avval yuklangan fayllar o'tkazib yuboriladi (dasturni qayta ishga
     tushirsangiz, to'xtagan joyidan davom etadi).

O'rnatish:
  pip install playwright pandas openpyxl
  python -m playwright install chrome   (agar kompyuterda Chrome bo'lmasa)

Ishga tushirish:
  python eshahodatnoma.py
"""

import ctypes
import sys
import time
import traceback
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

TAB = "2026"
STATUS = "Xulosa berildi"

KIRISH_KUTISH = 15 * 60  # login uchun beriladigan vaqt, soniya

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


def sahifa_belgisi(page):
    """Joriy sahifadagi birinchi ariza raqami: sahifa almashganini aniqlash uchun."""
    birinchi = page.locator(f"{GRID} tbody tr.rz-data-row a[href]").first
    return birinchi.inner_text() if birinchi.count() else ""


# Jadval yuklanib bo'lganini bildiradi: qatorlar bor yoki "yozuv yo'q" xabari chiqqan.
JADVAL_TAYYOR = """
grid => {
    const g = document.querySelector(grid);
    if (!g) return false;
    if (g.querySelector('tbody tr.rz-data-row')) return true;
    return !!g.querySelector('.rz-datatable-emptymessage');
}
"""


def qadam(matn):
    print(f"[{datetime.now():%H:%M:%S}] {matn}", flush=True)


def quickedit_ochirish():
    """Windows konsolida sichqoncha bosilsa dastur muzlab qolmasligi uchun QuickEdit o'chiriladi."""
    if sys.platform != "win32":
        return
    try:
        k32 = ctypes.windll.kernel32
        handle = k32.GetStdHandle(-10)  # STD_INPUT_HANDLE
        rejim = ctypes.c_uint()
        if k32.GetConsoleMode(handle, ctypes.byref(rejim)):
            # ENABLE_QUICK_EDIT_MODE (0x40) olib tashlanadi, ENABLE_EXTENDED_FLAGS (0x80) qo'yiladi
            k32.SetConsoleMode(handle, (rejim.value & ~0x40) | 0x80)
    except Exception:
        pass


def sayt_tabi(context):
    """Brauzerdagi e-shahodatnoma tabi (profil eski tablarni tiklagan bo'lishi mumkin)."""
    for tab in reversed(context.pages):
        if "e-shahodatnoma.uz" in tab.url:
            return tab
    return context.pages[-1] if context.pages else context.new_page()


def kirishni_kutish(context):
    """Foydalanuvchi saytga kirguncha kutadi: sahifada chiqish (logout) havolasi paydo bo'lishi."""
    qadam("Brauzerda login va parolingizni kiriting. Kirganingizdan so'ng dastur o'zi davom etadi.")
    tugash = time.time() + KIRISH_KUTISH
    while time.time() < tugash:
        for tab in list(context.pages):
            try:
                if "e-shahodatnoma.uz" in tab.url and tab.locator('a[href*="logout"]').count():
                    qadam("Saytga kirildi.")
                    return tab
            except Exception:
                pass  # tab yopilgan yoki sahifa almashayotgan bo'lishi mumkin
        time.sleep(1)
    raise RuntimeError(f"{KIRISH_KUTISH // 60} daqiqa ichida saytga kirilmadi")


def dublikat_filtrlash(page, yil=TAB, status=STATUS):
    """Dublikat sahifasi -> yil tabi -> Status ustuni filtri."""
    qadam("Dublikat bo'limi ochilmoqda...")
    page.goto(SAYT)
    page.wait_for_selector(".rz-tabview-nav a[role=tab]")

    qadam(f"{yil} tabi tanlanmoqda...")
    tab = page.locator(".rz-tabview-nav a[role=tab]", has_text=yil).first
    tab.click()
    page.wait_for_function("el => el.getAttribute('aria-selected') === 'true'",
                           arg=tab.element_handle())
    page.wait_for_function(JADVAL_TAYYOR, arg=GRID)

    qadam(f"Status = {status} filtri qo'yilmoqda...")
    oldingi_jami = jami_yozuv(page)
    # Filtr qatoridagi birinchi dropdown - Status ustuni (ikkinchisi - Operator)
    page.locator(f"{GRID} .filter-dropdown").first.click()
    page.locator(f'li.rz-dropdown-item[aria-label="{status}"]:visible').first.click()
    try:
        page.wait_for_function(
            """([sel, eski]) => {
                const el = document.querySelector(sel);
                return !el || el.textContent.replace(/\\D/g, '') !== String(eski);
            }""",
            arg=[".rz-paginator-summary", oldingi_jami],
            timeout=20_000,
        )
    except Exception:
        # Soni o'zgarmagan bo'lishi mumkin - baribir qatorlar Status bo'yicha tekshiriladi
        pass
    page.wait_for_function(JADVAL_TAYYOR, arg=GRID)
    qadam(f"{yil} tab, Status = {status}: {jami_yozuv(page) or 0} ta yozuv")


def jami_yozuv(page):
    summary = page.locator(".rz-paginator-summary").first
    if not summary.count():
        return None
    raqam = "".join(ch for ch in summary.inner_text() if ch.isdigit())
    return int(raqam) if raqam else None


def barcha_arizalarni_yigish(page, status=STATUS):
    if not page.locator(f"{GRID} tbody tr.rz-data-row").count():
        return []
    jami = jami_yozuv(page)
    qadam("Ro'yxat yig'ilmoqda...")

    natija, korilgan = [], set()
    sahifa_no = 1
    while True:
        qatorlar = page.locator(GRID).first.evaluate(QATORLARNI_OLISH)
        for q in qatorlar:
            kalit = q.get("Ariza raqami")
            # Filtr qo'yilmay qolgan bo'lsa ham faqat kerakli statusdagilar olinadi
            if q.get("Status") != status:
                continue
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


def ariza_yili(ariza):
    """Ariza sanasidan ("Sana / vaqt": 2026-10-30 13:46:36) yilni oladi."""
    yil = ariza.get("Sana / vaqt", "")[:4]
    return yil if yil.isdigit() else "sanasiz"


def excelga_saqlash(arizalar):
    NATIJALAR.mkdir(exist_ok=True)
    fayl = NATIJALAR / f"arizalar_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    df = pd.DataFrame(arizalar)
    df.insert(0, "Yil", [ariza_yili(a) for a in arizalar])

    with pd.ExcelWriter(fayl) as writer:
        df.to_excel(writer, sheet_name="Arizalar", index=False)
        for ustun in ["Yil", "Status", "Operator", "Tuman", "Maktab", "Muddat"]:
            if ustun in df:
                hisobot = (df[ustun].replace("", "(ko'rsatilmagan)")
                           .value_counts().rename_axis(ustun).reset_index(name="Soni"))
                hisobot.to_excel(writer, sheet_name=f"{ustun} bo'yicha"[:31], index=False)

    print(f"{len(df)} ta ariza saqlandi: {fayl}")
    print("\nYil bo'yicha:")
    print(df["Yil"].value_counts().sort_index().to_string())


def xulosa_faylini_yuklash(page, ariza_raqami, havola, papka):
    """Ariza sahifasidagi eng oxirgi "Xulosa berildi" bandining faylini papkaga yuklaydi."""
    if list(papka.glob(f"{ariza_raqami}.*")):
        return "oldin yuklangan"
    # Avvalgi versiya fayllarni yil papkasisiz saqlagan - ularni yil papkasiga ko'chiramiz
    eski = list(XULOSALAR.glob(f"{ariza_raqami}.*"))
    if eski:
        eski[0].rename(papka / eski[0].name)
        return "oldin yuklangan (yil papkasiga ko'chirildi)"

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
    manzil = papka / f"{ariza_raqami}{kengaytma}"

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
    qadam(f"Xulosa fayllari yuklanmoqda ({len(arizalar)} ta ariza)...")
    hisobot = []
    for i, ariza in enumerate(arizalar, 1):
        raqam, havola, yil = ariza.get("Ariza raqami"), ariza.get("Havola"), ariza_yili(ariza)
        papka = XULOSALAR / yil
        papka.mkdir(parents=True, exist_ok=True)
        try:
            holat = xulosa_faylini_yuklash(page, raqam, havola, papka)
        except Exception as xato:
            holat = f"xatolik: {str(xato).splitlines()[0]}"
        print(f"[{i}/{len(arizalar)}] {yil}/{raqam}: {holat}")
        hisobot.append({"Yil": yil, "Ariza raqami": raqam, "Holat": holat})

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


def xatoni_yozish(xato):
    """Xato ekranga va natijalar/xatolar.log fayliga yoziladi (tahlil uchun)."""
    print(f"\nXatolik: {xato}")
    try:
        NATIJALAR.mkdir(exist_ok=True)
        with open(NATIJALAR / "xatolar.log", "a", encoding="utf-8") as f:
            f.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n{traceback.format_exc()}")
        print(f"Batafsil: {NATIJALAR / 'xatolar.log'}")
    except Exception:
        pass


def main():
    quickedit_ochirish()
    qadam("Brauzer ochilmoqda...")
    with sync_playwright() as p:
        context = brauzerni_ochish(p)
        context.set_default_timeout(60_000)
        page = sayt_tabi(context)
        page.goto(SAYT)
        try:
            page = kirishni_kutish(context)
            dublikat_filtrlash(page)
            arizalar = barcha_arizalarni_yigish(page)
            if not arizalar:
                raise RuntimeError(f"{TAB} tabida '{STATUS}' arizasi topilmadi")
            excelga_saqlash(arizalar)
            xulosalarni_yuklash(page, arizalar)
            qadam("Tayyor.")
        except Exception as xato:
            xatoni_yozish(xato)
        input("\nYopish uchun Enter bosing... ")

        context.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as xato:
        # exe oynasi xatoni ko'rsatmasdan yopilib qolmasligi uchun
        xatoni_yozish(xato)
        input("Yopish uchun Enter bosing... ")
