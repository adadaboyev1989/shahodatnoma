# e-shahodatnoma arizalarini yig'ish

`e-shahodatnoma.uz` saytidagi **Arizalar** (`/duplicates`) jadvalini barcha sahifalari bilan yig'ib, Excel faylga saqlaydigan dastur.

## O'rnatish

```bash
pip install -r requirements.txt
python -m playwright install chrome   # kompyuterda Chrome bo'lmasa
```

## Ishlatish

```bash
python eshahodatnoma.py
```

1. Chrome ochiladi. Login va parolni o'zingiz kiritasiz (dastur parolni ko'rmaydi).
2. Kirgach terminalda Enter bosasiz. Dastur o'zi **Dublikat → 2026 → Status = Xulosa berildi** filtrini qo'yadi.
3. Filtrlangan jadvalning barcha sahifalari varaqlanadi.
4. Natija `natijalar/arizalar_YYYYMMDD_HHMMSS.xlsx` fayliga saqlanadi:
   - **Arizalar**: ariza raqami, havola, PINFL, status, operator, sana, muddat, F.I.O., viloyat, tuman, maktab
   - **Status / Operator / Tuman / Maktab / Muddat bo'yicha** hisobot varaqlari

`chrome_profil/` papkasida kirish sessiyasi saqlanadi, shuning uchun keyingi safar qayta kirish shart bo'lmasligi mumkin.
Bu papka ham, `natijalar/` ham shaxsiy ma'lumotlarni saqlaydi va `.gitignore` orqali repoga qo'shilmaydi.
