# calc_model — Metrology Core EURACHEM

Инструмент расчёта неопределённости калибровки по EURACHEM/CITAC QUAM, ISO 11095, GUM (JCGM 100/101).
Обратное предсказание концентрации по калибровочной прямой: ОМНК и взвешенная регрессия (WLS),
полная ковариационная матрица, Welch–Satterthwaite, Monte Carlo, бюджет неопределённости,
автоматическая валидация (20 проверок) и финальный отчёт в Word.

## Возможности

- **ОМНК + WLS** (веса 1, 1/x, 1/x²) с корректным исключением x=0 (TD-01)
- **Обратное предсказание**: x_pred = (y_obs − b0)/b1, полная интервальная форма u²(x_pred)
- **Ковариационная матрица** (XᵀWX)⁻¹, коэффициент дисперсии D = φ_w/s²_ОМНК
- **Welch–Satterthwaite** ν_eff, k по Стьюденту (TINV двусторонний)
- **Back-calculation** стандартов (ISO 11095, допуск ±15%), LOD/LOQ, t-тест бланка (b0)
- **Нормальность остатков** (SKEW/KURT), стандартизированные остатки, F-тест OLS vs WLS
- **Бюджет неопределённости**: Recovery, разбавление, навеска/объём/неоднородность, Pareto-диаграмма
- **Monte Carlo** (JCGM 101, N=10⁵) — независимая проверка GUM-аппроксимации
- **Автовалидация**: лист «Валидация» — 20 проверок + вердикт «КАЛИБРОВКА ПРИНЯТА/ОТКЛОНЕНА»
- **Отчёт Word**: титул, результат x±U (округление GUM 7.2.6), уравнения, бюджет, графики, подписи
- **Портативный exe**: работает с флешки на любом Windows-компе **без установки Python и Excel**

## Быстрый старт

### Вариант 1 — портативный exe (рекомендуется, без Python)

1. Скопируйте `dist/cal_app.exe` на флешку.
2. Дважды кликните или в cmd:
   ```
   cal_app.exe run            → полный цикл: расчёт → валидация → Monte Carlo → отчёт Word
   cal_app.exe calc           → только расчёт (создаёт Metrology_Core_EURACHEM_v4.xlsx)
   cal_app.exe mc 100000      → Monte Carlo, 100000 итераций
   cal_app.exe report         → отчёт Word (после расчёта)
   ```
3. В папке exe появятся: `Metrology_Core_EURACHEM_v4.xlsx`, `result.json`, `Отчёт_по_калибровке.docx`.
4. Правьте данные в зелёных ячейках xlsx и перезапустите `cal_app.exe run`.

### Вариант 2 — из исходников (Python)

Требуется Python 3.10+, пакеты: `openpyxl`, `numpy`, `scipy`, `matplotlib`, `python-docx`.
Опционально (для Excel-ветки с формулами): `pywin32` + установленный Microsoft Excel.

```
pip install openpyxl numpy scipy matplotlib python-docx pywin32
python cal_app.py run            # автономный полный цикл
python cal_v4.py                 # генерация xlsx с формулами (нужен Excel COM)
python validate_v4.py            # кросс-валидация numpy+scipy против Excel
python lib_validate.py           # проверка библиотеками (scipy/statsmodels)
python mc_v4.py 100000           # Monte Carlo
python report_v4.py              # отчёт Word из xlsx
```

## Сборка exe

```
pip install pyinstaller
python -m PyInstaller --onefile --console --name cal_app --clean --noconfirm cal_app.py \
  --hidden-import scipy --hidden-import matplotlib --hidden-import win32com \
  --collect-submodules scipy
```
Результат: `dist/cal_app.exe` (~100 МБ: scipy + matplotlib). Для уменьшения размера можно
исключить `win32com` (тогда только Python/MC-ветка).

## Структура проекта

| Файл | Назначение |
|------|-----------|
| `cal_app.py` | Единый CLI-конвейер (calc/validate/mc/report/run), точка входа для exe |
| `cal_v4.py` | Генератор xlsx с формулами (Excel-ветка, 15 листов) |
| `calc_engine.py` | Чистый Python-расчёт (без Excel COM) — автономная ветка |
| `validate_v4.py` | Кросс-валидация numpy+scipy против Excel (COM) |
| `validate_auto.py` | Автономная валидация (numpy, без Excel) |
| `lib_validate.py` | Проверка через scipy.stats.linregress / statsmodels.WLS |
| `mc_v4.py` | Monte Carlo по JCGM 101 (сравнение GUM vs MC) |
| `report_v4.py` | Отчёт Word |
| `test_td07.py` | Тест динамических диапазонов (TD-07) |
| `tracker.md` | Журнал версий |
| `techdebt.md` | Технические долги и план v5 |

## Ключевые результаты (демо-данные)

- x_pred = **0.544 ± 0.017** (k=2.006, ν_eff=52.9, P=0.95), LOD=0.0086, LOQ=0.0261
- ОМНК: b1=0.2737, b0=0.00125, R²=0.99996; WLS(1/x²): b1_w=0.2742
- **GUM vs MC**: u_GUM=0.00846 vs u_MC=0.00929 (расхождение 10.2%, JCGM 101 прим. 7 —
  при нелинейной обратной калибровке GUM занижает u, рекомендуется MC)
- Валидация: **20/20 PASS**, вердикт «КАЛИБРОВКА ПРИНЯТА»

## Лицензия

Для внутреннего использования в лаборатории. Не является заменой официальной
программы обеспечения качества — применяйте вместе с валидацией метода.