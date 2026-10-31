# e-shahodatnoma arizalarini yig'ish

`e-shahodatnoma.uz` saytidagi **Arizalar** (`/duplicates`) jadvalini barcha sahifalari bilan yig'ib, Excel faylga saqlaydigan dastur.

## Tayyor exe (Windows)

GitHub'da **Actions → Windows exe** bo'limidagi oxirgi muvaffaqiyatli yig'ishni oching va pastdagi
**Artifacts** qismidan `eshahodatnoma-exe` ni yuklab oling (zip ichida `eshahodatnoma.exe`).
Python o'rnatish shart emas, faqat kompyuterda Chrome yoki Edge bo'lishi kerak.
`natijalar/` va `chrome_profil/` papkalari exe yonida yaratiladi, shuning uchun uni alohida papkaga qo'ying.

## Python orqali o'rnatish

```bash
pip install -r requirements.txt
python -m playwright install chrome   # kompyuterda Chrome bo'lmasa
```

## Ishlatish

```bash
python eshahodatnoma.py
```

1. Chrome ochiladi. Login va parolni o'zingiz kiritasiz (dastur parolni ko'rmaydi).
2. Kirganingizni dastur o'zi sezadi (Enter bosish shart emas) va **Dublikat → 2026 → Status = Xulosa berildi** filtrini qo'yadi.
3. Filtrlangan jadvalning barcha sahifalari varaqlanadi. 
4. Natija `natijalar/arizalar_YYYYMMDD_HHMMSS.xlsx` fayliga saqlanadi:
   - **Arizalar**: ariza raqami, havola, PINFL, status, operator, sana, muddat, F.I.O., viloyat, tuman, maktab
   - **Status / Operator / Tuman / Maktab / Muddat bo'yicha** hisobot varaqlari
5. So'ng har bir arizaga kirib, **Xulosa** blokidagi vaqti eng so'nggi «Xulosa berildi» bandining faylini yuklaydi.
   Fayllar «Sana / vaqt» ustunidagi yil bo'yicha ajratiladi (2026 tabida 2025-yil sanali arizalar ham bor):
   `natijalar/xulosalar/2026/<ariza raqami>.pdf`, `natijalar/xulosalar/2025/<ariza raqami>.pdf` va hokazo.
   Avval yuklangan fayllar o'tkazib yuboriladi (eski versiyada yil papkasisiz saqlanganlari yil papkasiga ko'chiriladi),
   shuning uchun dastur to'xtab qolsa, qayta ishga tushirish kifoya.

Xatolik bo'lsa, batafsil ma'lumot `natijalar/xatolar.log` fayliga yoziladi.

`chrome_profil/` papkasida kirish sessiyasi saqlanadi, shuning uchun keyingi safar qayta kirish shart bo'lmasligi mumkin.
Bu papka ham, `natijalar/` ham shaxsiy ma'lumotlarni saqlaydi va `.gitignore` orqali repoga qo'shilmaydi.
