# reward_matrix — бизнес-логика и спецификация инструмента

Документ из трёх частей:
- **Часть A — Доменные знания.** Что такое reward-матрица и как она работает. Не привязано к конкретному скрипту или интерфейсу.
- **Часть B — Legacy-скрипты.** Краткое описание того, что сейчас лежит в `legacy/` и зачем.
- **Часть C — Новый инструмент (MVP Phase 1).** Полная спецификация того, что надо построить: модель данных, SQLite-схема, экраны, контракты.

---

# ЧАСТЬ A. ДОМЕННЫЕ ЗНАНИЯ

## A.1 Что такое reward-матрица

Матрица — это набор правил, по которым банк платит компании комиссионное вознаграждение (**КВ**, в коде — `reward`) за выдачу POS-кредита. КВ зависит от параметров сделки: банк, тип обработки (`inside`/`outside`), тариф, ставка, срок, сумма, банковский продукт (`stock`).

**Принцип расчёта:**
- В матрице есть **группы условий** (AND внутри группы)
- Если все условия группы выполнены → группа «сработала»
- Между группами — OR, из всех сработавших берётся **минимальный** `reward_value`
- Если ни одна группа не прошла — фиксируется причина в `matrix_status`

## A.2 Структура правила

Каждое правило — строка с полями:

| поле | тип | пример |
|------|-----|--------|
| `bank_name` | str | `"otp"` |
| `proc_type` | str | `"inside"` / `"outside"` |
| `tariff_code` | str | `"general"` |
| `group` | str | `"grp_1"` (любое имя, группирует условия по AND) |
| `parameter` | str | `"rate"`, `"term"`, `"stock"`, `"sumcredit"`, ... |
| `condition_type` | str | `=`, `!=`, `<`, `<=`, `>`, `>=`, `between`, `in`, `not_in`, `like`, `not_like` |
| `value` | str | `"10;20"` для between, `"%Выгодный%"` для like |
| `reward_value` | float | процент КВ (доля от суммы кредита) |

## A.3 Операторы condition_type

- **Базовые:** `=`/`==`, `!=`/`<>`, `<`, `<=`, `>`, `>=`
- **Диапазон:** `between` с форматом `"a;b"` (по умолчанию `[a;b)`), можно явно `[a;b]`, `(a;b)` и т.д.
- **Списки:** `in` / `not_in` — значения через `;`
- **SQL-LIKE:** `like` / `not_like` — шаблоны с `%` (любое количество символов) и `_` (один символ), несколько шаблонов через `;`

## A.4 Привязка договора к тарифу (`tariff_agreements`)

Входящая сделка из DWH содержит `agreement_name` — название конкретного договора с банком. В матрице правила привязаны не к договорам, а к `tariff_code` (условный код тарифа). Таблица `tariff_agreements` делает маппинг:

```
(bank_name, proc_type, agreement_name) → tariff_code
```

Один `tariff_code` покрывает несколько договоров. Один договор — строго один `tariff_code`.

## A.5 `stock` и `stock_code` — банковский продукт

`stock` — банковский продукт, который банк публикует под акцию или конкретный сегмент («Выгодный 10-24», «Рассрочка 0-0-24» и т.п.). `stock_code` — идентификатор этого продукта в DWH. Внутри одного `tariff_code` может быть несколько `stock`-ов, и правила матрицы часто завязаны именно на `stock` (через `like` с маской или `in` со списком).

Важно: `stock` — это отдельный параметр условия в группе, а не часть ключа `(bank, proc, tariff)`. Не путать с `idstock` (суррогат из DWH).

## A.6 Fallback — это компромисс, а не фича

Когда сделка не покрыта матрицей, действует fallback. Возможны два случая:

| статус | когда | что ставится в `reward_matrix` |
|--------|-------|--------------------------------|
| `NO_BANK` | банка нет в `tariff_agreements` | `real_reward` (фактический КВ из DWH) |
| `NO_PROC_TYPE` | банк есть, типа обработки нет | `real_reward` |
| `NO_AGREEMENT` | банк+proc есть, договор не смаппирован | `real_reward` |
| `NO_<parameter>` | договор смаппирован, но ни одна группа не прошла (например, ставка не попала в диапазон) | `#N/A` |
| `OK` | всё прошло | `min(reward_value)` среди прошедших групп |

**Это компромисс.** Fallback на `real_reward` означает, что для этих сделок сценарий ничего не меняет — они всегда считаются как есть в жизни. Чем больше дыр в `tariff_agreements`, тем меньше реальных плеч у сценария влиять на итог.

**`#N/A`** — ещё бо́льшая дыра: сделка выбывает из численного КВ по сценарию. Такие сделки в отчёте должны агрегироваться **отдельной категорией «не покрыто матрицей»** с их фактическим `real_reward`, чтобы сумма по сценарию сходилась с реальностью.

## A.7 Индикаторы качества матрицы

В отчёте по каждому сценарию всегда показывать:

- доля сделок со статусом `NO_BANK`, `NO_PROC_TYPE`, `NO_AGREEMENT`, `NO_*` (в штуках и в % от объёма выдач)
- доля сделок со статусом `#N/A`
- суммарная разница `reward_matrix − real_reward` с разбивкой по статусам

Если доля покрытия падает → матрица теряет способность отличать сценарии. Это сигнал «пора чинить `tariff_agreements` и правила», а не «сценарий плохой».

## A.8 Колода матриц и именование сценариев

Сценарий = набор правил, применяемый ко всем сделкам месяца. Разные сценарии нужны для сравнения: «как есть сейчас», «как будет, если банк А даст такие условия», «гипотеза: а что если ввести новый грейд».

**Типы сценариев:**

| префикс | смысл | пример |
|---------|-------|--------|
| `current_` | действующие условия на период | `current_2026-04` |
| `offer_` | предложение банка/контрагента | `offer_alfa_2026-05` |
| `hypothesis_` | внутренняя гипотеза «что если» | `hypothesis_grade500m_2026-04` |
| `what_if_` | комбинированный what-if | `what_if_rate_down_2026-04` |

**Правила нейминга:**
- латиница + цифры + `_` + `-`, без пробелов и кириллицы
- ограничение 31 символ (наследие Excel, пусть будет единообразно)
- формат: `<type>_<subject>_<YYYY-MM>`

## A.9 Приоритизация банков (`bank_priority`)

Отдельная таблица со справочной ролью. Два независимых измерения в одной таблице:

| поле | тип | смысл | где используется |
|------|-----|-------|------------------|
| `bank_name` | str | банк | ключ |
| `proc_type` | str | тип обработки | ключ |
| `tariff_code` | str (NULL ok) | тариф, опционально | ключ |
| `month` | str `YYYY-MM` | период приоритета | ключ |
| `priority` | int | ранг (1 — первым слать заявки) | оперативный контур |
| `target_share_pct` | float (0..100) | целевая доля объёма | аналитический контур, what-if |

Обе колонки опциональны: можно заполнить только `priority`, только `target_share_pct`, или обе. Если для текущего месяца нет записи — просто нет приоритезации.

---

# ЧАСТЬ B. LEGACY-СКРИПТЫ

В `legacy/` лежат два рабочих Python-скрипта. Их **нельзя переписывать** — логика расчёта в них устоялась и проверена на реальных данных. В новом инструменте они используются как библиотеки (импортируем функции, не правим).

## B.1 `legacy/matrix_reward.py` — аудитный расчёт

- Один лист матрицы из `audit_deck.xlsm` → один прогон
- Читает месяц из DWH (`mart_bi.excel_page_plan_kv_2`)
- Пишет результат в `matrix_reward.xlsx`, один лист `Data`
- Используется для аудита: взять фактический месяц и проверить, совпал ли расчётный КВ с банковским

## B.2 `legacy/reward_matrix.py` — аналитический расчёт

- N листов-сценариев из того же `audit_deck.xlsm` → один прогон
- Тот же источник данных
- Пишет в `reward_compression.xlsx`: лист `Data` + лист `ScenarioResults` (длинная таблица со столбцом `scenario_name`)
- Используется для сравнения сценариев: what-if, оценка предложений банков

## B.3 Ключевые переиспользуемые функции

- `clean_text`, `norm_proc`, `norm_agr` — нормализация
- `eval_cond` — проверка одного условия
- `evaluate_groups` — проверка всех групп для одной строки данных
- `build_rules_index` — из DataFrame правил строит индекс `key → groups`
- `build_agreement_index` — из DataFrame `tariff_agreements` строит `bank → proc → agreement → tariff_code`
- `load_month_from_dwh` — чтение DWH через pyodbc

Они работают с pandas DataFrame и словарями. Новый инструмент должен подавать данные в том же формате — это позволит переиспользовать функции без рефакторинга.

---

# ЧАСТЬ C. НОВЫЙ ИНСТРУМЕНТ (MVP Phase 1)

## C.1 Назначение

Десктоп-подобная программа через браузер (Streamlit): пользователь импортирует `audit_deck.xlsm`, получает набор **атомарных матриц-шаблонов** (по одной на `bank+proc+tariff`), собирает из них **сценарии**, запускает расчёт в двух режимах (`audit` или `analytics`), получает xlsx-отчёт.

## C.2 Технологический стек

| слой | выбор | обоснование |
|------|-------|-------------|
| UI | Streamlit | Python-only, браузерный клиент, быстрый старт |
| Хранилище | SQLite + SQLAlchemy | локально у каждого пользователя, zero-config |
| Миграции | Alembic | управляемые изменения схемы |
| Расчёт | `legacy/reward_matrix.py` как библиотека | не переписывать логику |
| Чтение xlsm | openpyxl | уже зависимость |
| DWH | pyodbc через DSN | как сейчас |
| Экспорт | openpyxl | как сейчас |

Хранилище **локальное у каждого пользователя** (отдельный `app.db`). Никакой централизации на MVP.

## C.3 Модель данных

Четыре основные сущности:

```
┌─────────────────┐     ┌─────────────────────┐
│    Template     │─────│  template_rule      │
│  (атом-матрица) │  1:N│  (условие группы)   │
└────────┬────────┘     └─────────────────────┘
         │ 1:N
         ├─────► template_agreement (договора)
         │
         │ M:N через scenario_template
         ▼
┌─────────────────┐
│    Scenario     │
│  (композиция)   │
└────────┬────────┘
         │ 1:N
         ▼
┌─────────────────┐
│      Run        │
│   (прогон)      │
└─────────────────┘

┌─────────────────┐   (отдельная таблица, справочник)
│  BankPriority   │
└─────────────────┘
```

## C.4 Схема SQLite

```sql
-- атом-матрица: один банк × один proc_type × один tariff_code
CREATE TABLE template (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL UNIQUE,        -- 'otp_outside_general__tm_2026-02'
    bank_name       TEXT NOT NULL,
    proc_type       TEXT NOT NULL,               -- 'inside' | 'outside'
    tariff_code     TEXT NOT NULL,
    description     TEXT,                        -- заметка (например, 'грейд 100 млн')
    is_favorite     INTEGER NOT NULL DEFAULT 0,
    is_archived     INTEGER NOT NULL DEFAULT 0,
    created_at      TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at      TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- правила конкретного шаблона: плоская форма (как в xlsm)
CREATE TABLE template_rule (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id     INTEGER NOT NULL REFERENCES template(id) ON DELETE CASCADE,
    group_name      TEXT NOT NULL,               -- 'grp_1'
    parameter       TEXT NOT NULL,               -- 'rate', 'term', 'stock', ...
    condition_type  TEXT NOT NULL,               -- '=', 'between', 'like', ...
    value           TEXT NOT NULL,
    reward_value    REAL                         -- nullable: reward есть на первой строке группы
);
CREATE INDEX idx_rule_template ON template_rule(template_id);
CREATE INDEX idx_rule_group ON template_rule(template_id, group_name);

-- связь шаблона с договорами (замена tariff_agreements из xlsm)
CREATE TABLE template_agreement (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id     INTEGER NOT NULL REFERENCES template(id) ON DELETE CASCADE,
    agreement_name  TEXT NOT NULL,
    UNIQUE (template_id, agreement_name)
);
CREATE INDEX idx_agr_template ON template_agreement(template_id);

-- сценарий: композиция шаблонов
CREATE TABLE scenario (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL UNIQUE,        -- 'offer_alfa_2026-05'
    description     TEXT,
    is_favorite     INTEGER NOT NULL DEFAULT 0,
    is_archived     INTEGER NOT NULL DEFAULT 0,
    created_at      TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at      TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- M:N связь сценария и шаблонов
CREATE TABLE scenario_template (
    scenario_id     INTEGER NOT NULL REFERENCES scenario(id) ON DELETE CASCADE,
    template_id     INTEGER NOT NULL REFERENCES template(id) ON DELETE RESTRICT,
    PRIMARY KEY (scenario_id, template_id)
);

-- запуск расчёта
CREATE TABLE run (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    scenario_id     INTEGER NOT NULL REFERENCES scenario(id),
    scenario_name   TEXT NOT NULL,               -- снапшот на момент запуска (для истории)
    mode            TEXT NOT NULL,               -- 'audit' | 'analytics'
    month           TEXT NOT NULL,               -- 'YYYY-MM'
    output_path     TEXT NOT NULL,
    status          TEXT NOT NULL,               -- 'running' | 'done' | 'failed'
    error_message   TEXT,
    started_at      TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    finished_at     TEXT
);
CREATE INDEX idx_run_scenario ON run(scenario_id);

-- приоритизация банков
CREATE TABLE bank_priority (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    bank_name           TEXT NOT NULL,
    proc_type           TEXT NOT NULL,
    tariff_code         TEXT,                    -- NULL = общий приоритет банка
    month               TEXT NOT NULL,           -- 'YYYY-MM'
    priority            INTEGER,
    target_share_pct    REAL,
    UNIQUE (bank_name, proc_type, tariff_code, month)
);
```

## C.5 Экраны UI

Минимальный набор — 4 страницы в Streamlit sidebar:

### C.5.1 `Templates`

- Таблица: `name`, `bank_name`, `proc_type`, `tariff_code`, `is_favorite`, `updated_at`
- Фильтры: по банку, proc_type, флаг «только избранное», «не архив»
- Действия: создать (форма), редактировать (форма), клонировать, архивировать, восстановить
- Форма редактирования: шапка шаблона + таблица правил с inline-редактированием (добавить/удалить строку) + таблица договоров

### C.5.2 `Scenarios`

- Таблица: `name`, кол-во включённых шаблонов, `is_favorite`, `updated_at`
- Фильтры: те же что у Templates
- Действия: создать, редактировать состав (добавить/убрать Template), клонировать, архивировать
- Форма редактирования: шапка + двухпанельный компоновщик (слева доступные Templates с поиском, справа включённые)

### C.5.3 `Runs`

- Таблица: `scenario_name`, `mode`, `month`, `status`, `started_at`, `output_path` со ссылкой на файл
- Кнопка «Новый прогон»: выбор сценария → выбор режима (`audit` / `analytics`) → выбор месяца (latest / YYYY-MM) → выбор пути для xlsx → кнопка «Запустить»
- Запуск синхронный (без фонового воркера на MVP)
- Прогресс через `st.status` / `st.progress`

### C.5.4 `Import`

- Загрузка xlsx (тот самый `audit_deck.xlsm` или любой с такой же структурой)
- Автоопределение листов-матриц (проверка наличия колонок `bank_name`, `proc_type`, `tariff_code`, `group`, `parameter`, `condition_type`, `value`, `reward_value`)
- Выбор листов для импорта (галочки)
- Превью: сколько шаблонов получится, есть ли конфликты имён
- Стратегия конфликта: `skip` / `overwrite` / `rename with suffix`
- Кнопка «Импортировать»

## C.6 Контракт импортёра

**Вход:** путь к xlsx + список имён листов + стратегия конфликта.

**Что делает на каждом выбранном листе:**

1. Читает лист как DataFrame
2. Читает лист `tariff_agreements` из того же xlsx (если есть)
3. Группирует правила по `(bank_name, proc_type, tariff_code)` → получает N потенциальных Template'ов
4. Для каждой группы:
   - генерирует имя: `{bank}_{proc}_{tariff}__{sheet_name}`, нормализация: lower, без пробелов/спецсимволов
   - создаёт Template (с описанием «импортировано из листа `X` `YYYY-MM-DD HH:MM`»)
   - создаёт Template_rule для каждой строки правила
   - из `tariff_agreements` берёт все `agreement_name` для этого `(bank, proc, tariff)` → создаёт Template_agreement
5. Создаёт Scenario с именем = имя листа и прикрепляет к нему все созданные/найденные Template'ы

**Конфликты имён:**
- `skip` — шаблон с таким именем уже есть → пропустить (к сценарию всё равно прицепить существующий)
- `overwrite` — удалить старый шаблон и создать заново
- `rename with suffix` — дописать `_v2`, `_v3` ... к имени нового

**На выходе:** лог операций (что создано, пропущено, переименовано).

## C.7 Контракт раннера

**Вход:** `scenario_id`, `mode` (`audit` | `analytics`), `month` (`YYYY-MM`), путь для xlsx.

**Что делает:**

1. Поднимает шаблоны сценария из SQLite
2. Собирает DataFrame правил в формате, совместимом с `legacy/reward_matrix.py::build_rules_index` (колонки: `bank_name`, `proc_type`, `tariff_code`, `group`, `parameter`, `condition_type`, `value`, `reward_value`)
3. Собирает DataFrame договоров в формате для `build_agreement_index` (колонки: `bank_name`, `proc_type`, `agreement_name`, `tariff_code`)
4. Вызывает `load_month_from_dwh` — получает данные за месяц
5. Прогоняет через `evaluate_groups` (одну сделку за раз, как в legacy)
6. Собирает итоговый DataFrame со статусами и расчётным `reward_matrix`
7. Пишет xlsx:
   - **mode=audit** — один лист `Data` с исходными данными + колонки `reward_matrix`, `matrix_status`, `matrix_group`, `reward_matrix_amount`
   - **mode=analytics** — два листа: `Data` (исходник) и `ScenarioResults` (длинная таблица с `scenario_name`)
8. Обновляет `run.status`, `run.finished_at`

**Совместимость:** результат должен **побайтно** совпадать с тем, что даёт `legacy/reward_matrix.py` для одного и того же месяца и одного и того же сценария. Это проверяется тестом.

## C.8 Чего в MVP Phase 1 НЕ будет

Явно откладывается на следующие фазы:

- визард импорта из **банковских** 2D-таблиц (не xlsm-deck, а оригиналы от банков)
- сравнение сценариев между собой (diff)
- версионирование шаблонов (Published/Current/Change log из общей архитектуры)
- аудит-лог действий пользователя
- авторизация, профили, мультипользовательский доступ
- UI-редактирование `bank_priority` (на MVP — заливка через SQL или CSV-импорт)
- фоновые задачи, очередь, прогресс по шагам

Это важные вещи, но попытка впихнуть их в Phase 1 превратит недельную работу в месячную. MVP должен заменить два текущих скрипта — не больше.

## C.9 Критерии приёмки MVP

Явный чек-лист, по которому проверяется, что Phase 1 сделана:

1. Приложение запускается командой `streamlit run app.py`
2. При первом запуске создаётся `app.db`, применяются миграции
3. Импорт `audit_deck.xlsm`: выбираем 2 листа → создаются Template'ы и Scenario'и без ошибок
4. На странице Templates видим импортированные шаблоны, можем открыть форму, отредактировать правило, сохранить
5. На странице Scenarios видим импортированные сценарии, можем собрать новый из существующих шаблонов
6. На странице Runs запускаем сценарий в режимах `audit` и `analytics`, xlsx создаётся по указанному пути
7. **Регрессионный тест:** результат раннера побайтно совпадает с `legacy/reward_matrix.py` для того же сценария и месяца
8. Все действия логируются в консоль с меткой времени и уровнем (INFO/WARNING/ERROR)
