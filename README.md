### 📘 `README.md` — SpellCurrencyUniversal
![Version](https://img.shields.io/badge/version-v1.0.0-blue.svg)
![Supported Currencies](https://img.shields.io/badge/Supported%20Currencies-12-blue)
```markdown
# 💸 SpellCurrencyUniversal

A Unicode-aware, locale-sensitive Excel VBA engine that converts currency-formatted numbers into full words—with support for Indian and Western grouping systems, multi-currency detection, and smart metadata-driven extensibility.

---

## ✨ Features

- 🧠 Automatic detection of currency via Excel formatting, Unicode symbols, or ISO prefixes (₹, USD, €, ¥, etc.)
- 📐 Dual grouping styles: Indian (Lakh/Crore) and Western (Thousand/Million)
- 🔁 Modular architecture: easily add new currencies and styles
- 🧼 Cleans, parses, and validates even messy or copied currency text (like "USD 0.00", "€5 000,75")
- ❌ Gracefully handles invalid or non-numeric inputs
- 🔒 Password-locked VBA module for basic IP protection (optional)

---

## 📦 Installation

1. Download [`SpellCurrencyUniversal.xlam`](#) (Add-In format)
2. Open Excel
3. Go to: `File > Options > Add-Ins > Manage: Excel Add-ins > Browse`
4. Select the `.xlam` file and enable it
5. Use the function directly in any cell:

```excel
=SpellCurrencyUniversal(A1)
```

---

## 🛠 How to Add a New Currency

1. **Edit `GetCurrencyMetadataByISO(isoCode)`**:

   ```vba
   Case "XYZ"
       meta.Add "unit", "Zorbs"
       meta.Add "subunit", "Microzorbs"
       meta.Add "format", "Western"  ' or "Indian"
   ```

2. **Add detection in `DetectISOFromCell(cell)`**:

   ```vba
   If InStr(fmt, "xyz") > 0 Or InStr(cell.Text, "XYZ") > 0 Then
       DetectISOFromCell = "XYZ": Exit Function
   ```

3. **(Optional)** Add Unicode detection in `DetectISOFromDisplayText(cell)`:

   ```vba
   Case 8379 ' Unicode of symbol
       DetectISOFromDisplayText = "XYZ": Exit Function
   ```

---

## 🌐 Supported Currencies

Includes full detection and formatting for:

- INR – Indian Rupee (₹)
- USD – US Dollar ($)
- EUR – Euro (€)
- JPY – Japanese Yen (¥)
- GBP – British Pound (£)
- AED – UAE Dirham (د.إ)
- ILS – Israeli Shekel (₪)
- KRW – Korean Won (₩)
- RUB – Russian Ruble (₽)
- VND – Vietnamese Dong (₫)
- CHF – Swiss Franc
- TWD – New Taiwan Dollar (NT$)
- and others...

---

## 🧪 Testing

Comes with a test sheet (`/tests/TestCases.xlsx`) featuring 100+ real-world currency formats for validation. Includes:

- Symbol-prefixed and ISO-suffixed formats
- Decimal and zero-decimal currencies
- Invalid inputs and edge cases

---

## 🔐 Licensing & IP

This project is distributed as compiled `.xlam` with locked VBA modules. For inquiries about licensing, contribution, or commercial use, please contact **Kannan** (via Issues or Discussions tab).

---

## 🧙 About the Author

Crafted by [Kannan](#) — Excel wizard, modular design enthusiast, and seeker of universal linguistic and numerical harmony. Built with support from Microsoft Copilot.

```
