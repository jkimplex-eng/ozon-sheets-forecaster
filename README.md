# Ozon Sheets Forecaster

Google Sheets + Apps Script проект для сбора и анализа данных Ozon Seller API и Ozon Performance API.

Этот репозиторий создан специально для проекта Ozon. Не используйте для него `ROI_based_video_encoding` или другие Go/Python репозитории.

## Архитектура

- Google Sheets выступает как интерфейс и база таблиц.
- Apps Script (`Code.gs`) запускается из меню Google Таблицы.
- Google Cloud и `service-account.json` не нужны.
- Deploy / Развернуть не нужен для текущей версии, потому что это container-bound Apps Script.
- Ozon ключи хранятся только в `PropertiesService.getScriptProperties()`.

## Как установить Code.gs в Google Sheets

1. Открой рабочую Google Таблицу.
2. Нажми `Расширения -> Apps Script`.
3. Вставь содержимое `Code.gs`.
4. Нажми Save / Сохранить.
5. Вернись в таблицу и обнови страницу.
6. Должно появиться меню `Ozon`.

## Порядок запуска

1. `Ozon -> Создать вкладки`
2. `Ozon -> Настроить ключи Ozon`
3. `Ozon -> Проверить подключение Ozon`
4. `Ozon -> Обновить товары`
5. `Ozon -> Обновить остатки`
6. `Ozon -> Обновить последние 30 дней`
7. `Ozon -> Диагностика рекламных кампаний`
8. `Ozon -> Debug рекламы` или `Ozon -> Обновить статистику рекламы`
9. `Ozon -> Построить прогноз`
10. `Ozon -> Обновить Summary`

## Что уже должно работать в текущей версии Apps Script

По текущему состоянию проекта в Google Sheets уже подтверждено:

- Products подтягиваются.
- Stocks подтягиваются.
- Продажи и Finance подтягиваются.
- Ad_Campaigns подтягивается через Performance API `GET /api/client/campaign`.
- Summary строит базовый P&L с Unit_Costs.

## Что ещё нужно доделать

Основной незакрытый блок — фактическая статистика рекламы:

- `Ad_Campaigns` = справочник кампаний.
- `Raw_Ads_Daily` = фактические показы, клики, расходы, revenue attribution.
- Summary/P&L сможет учесть рекламу только после заполнения `Raw_Ads_Daily.spend`.

Следующая задача описана в `docs/codex_next_task_ads_debug.md`.

## Важно про Performance API

Использовать только новый host:

```text
https://api-performance.ozon.ru
```

Список кампаний:

```text
GET /api/client/campaign
```

Ответ приходит в поле `list`.

Если кампании есть в интерфейсе Ozon, но API показывает 0:

1. проверь, что Performance API key создан в том же рекламном аккаунте;
2. запусти `Ozon -> Диагностика рекламных кампаний`;
3. смотри вкладки `Ad_Campaigns` и `Diagnostics`.

## 429 rate limit

Если Performance API возвращает:

```text
Превышен лимит активных запросов (максимум 1)
```

нужны последовательные запросы, `Utilities.sleep(...)` и retry/backoff. Не запускать много Performance-запросов параллельно.

## P&L и себестоимость

Себестоимость хранится во вкладке `Unit_Costs`:

```text
offer_id | unit_cost | currency | updated_at | notes
```

P&L считает:

```text
Выручка gross
- Себестоимость товара
- Комиссия Ozon
- Логистика Ozon
- Возвраты / удержания
- Рекламные расходы
= Операционная прибыль
```

Если комиссии/логистики/рекламы нет, P&L должен показывать статус `Неполный`, а не молча считать отсутствующие расходы как 0.

## Безопасность

Никогда не коммитить и не выводить:

- `OZON_SELLER_API_KEY`
- `OZON_PERFORMANCE_CLIENT_SECRET`
- `access_token`
- `Authorization` header value
- любые реальные секреты

В diagnostics можно писать только `есть / нет`, HTTP code, короткий preview ответа без токенов.

## Как продолжать через Codex

Перед любой задачей Codex должен проверить, что он находится в правильном репозитории:

```bash
ls
rg --files | rg "Code.gs|README.md"
```

Если `Code.gs` отсутствует — остановиться. Это не тот репозиторий.

Если `Code.gs` пока placeholder, сначала скопируй текущий рабочий код из Google Apps Script в `Code.gs`, затем продолжай задачи Codex.
