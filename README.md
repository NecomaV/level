# level — генератор формул «Общая сводка все филиалы»

`app/overall_summary_autodiscovery.py` автогенерирует формулы листа
**«Общая сводка все филиалы»** в Google-таблице маркетинга на основе
discovery блоков в листах «Сводка …».

## Что делает

Для каждой строки-филиала в колонке `B` overall-листа находит соответствующий
лист «Сводка …», автоматически определяет колонки (бюджет, лиды, приходы,
записи, платники, продажи) и записывает формулы в `C:L` с учётом:

- режима периода (`$B$1`: «Весь период» / «Определённая дата» / «Диапазон дат»);
- конвертации валют:
  - Дубай / CocoAge → `$E$1` (USD→AED), всегда;
  - Ташкент → платники `$H$1` (USD→UZS) при `$G$1="USD"`, продажи **не делятся**;
  - KZ-филиалы и франшизы → `$F$1` (USD→KZT) при `$G$1="USD"`.

Обрабатываются обе таблицы — **«Филиал»** и **«Франшиза»** (у каждой свой «Итог:»).
Блок **«KPI»** и всё ниже него **не трогается**.

## Установка

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

Создай файл `.env` рядом с проектом:

```
GOOGLE_SERVICE_ACCOUNT_JSON=C:\path\to\service-account.json
```

Сервис-аккаунт должен иметь доступ к таблице (расшарь таблицу на email
сервис-аккаунта с правом редактирования).

## Запуск

ID таблицы — из URL: `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`.

```bash
# 1) сухой прогон — ничего не пишет, только показывает диапазоны и секции
python app/overall_summary_autodiscovery.py --spreadsheet_id <SPREADSHEET_ID> --dry_run

# 2) карта найденных блоков для ручной сверки (опционально)
python app/overall_summary_autodiscovery.py --spreadsheet_id <SPREADSHEET_ID> \
    --write_blockmap blockmap.json --only_blockmap

# 3) реальная запись формул
python app/overall_summary_autodiscovery.py --spreadsheet_id <SPREADSHEET_ID> --apply
```

### Параметры

| Флаг | По умолчанию | Описание |
|---|---|---|
| `--spreadsheet_id` | — (обязательный) | ID Google-таблицы |
| `--work_rows` | `4:33` | строки дат в листах «Сводка …» (июнь = 30 дней → `4:33`, февраль = 28 → `4:31`) |
| `--scan` | `B:ET` | диапазон колонок для discovery |
| `--overall_sheet` | `Общая сводка все филиалы` | имя overall-листа |
| `--overall_labels_range` | `B3:B60` | где искать метки филиалов |
| `--apply` | — | реально записать (без него — dry-run) |
| `--dry_run` | — | принудительный сухой прогон |
