UzWealth AI — n8n workflows
Целевой экземпляр n8n Cloud 2.44.5

Создано в рабочей области
1. F100 deterministic tests
https://dilbarb.app.n8n.cloud/workflow/ubNoxSevyePZyeHr
2. Backend integration draft
https://dilbarb.app.n8n.cloud/workflow/Nn235rfHspveHoSB
Оба workflow сохранены с active=false. Webhook и расписания отсутствуют.

Что работает
F100 выполнен в n8n. Execution ID 1: Succeeded, 1.849s.
Итоговый узел: tests_passed=true; accepted=1, duplicate=1, rejected=1.
Валидное событие: event_id=F100-001, capital_tiyin=10000000000,
horizon_months=12, effective_risk=moderate, confirmed=true.
Второе событие — его точная копия. Третье содержит capital_tiyin=-100.
Расчёт использует синтетические цены, целые слитки и BigInt.
Ожидаемый результат: депозит 6000000000, золото 2300000000,
BTC 1485148000, комиссия BTC 14851480, cash 200000520 тийинов.
BTC quantity=0.01485148. Итого 10000000000 тийинов (100 млн UZS).
Экранные доли net cost: deposit6009/gold2304/BTC1487/cash200 bps.
Прототип поддерживает только профиль F100 и не выдаёт реальных рекомендаций.
Set обнаруживает дубликаты внутри запуска; между запусками он не сохраняется.

Что ещё настроить для интеграции
Сам backend FastAPI/LangGraph нужно развернуть отдельно.
В Integration config заполнить backend_base_url (HTTPS origin) и event_id
(UUID бизнес события; тот же UUID при повторной попытке).
Idempotency-Key равен lowercase event_id без префикса.
В обоих HTTP Request привязать собственный n8n credential Simplified Custom Auth
с подходящим backend заголовком Authorization. Токены в код не помещаются.
Реальные credentials не созданы и не экспортированы.
Пустая конфигурация обрабатывается как CONFIG_REQUIRED до HTTP.
Интеграционный workflow не запускался; сеть и API end-to-end не проверены.
Черновик содержит POST generate, GET job, проверки ответов и общий deadline45s.
Redirects отключены; HTTP ошибки передаются в обработчик с прежним ключом.
Wait проверен в UI: 5 Seconds. В копии JSON это значение записано явно,
так как экспорт текущего экземпляра опускает параметры со значениями по умолчанию.
Production crypto остаётся выключенной; LLM не выполняет финансовую арифметику.

Проверки
Оба JSON проверены штатным структурным checker из n8n-workflow-check.
run_checks.mjs выполняет Code-узлы в локальном JS VM без HTTP и credentials:
валидный/невалидный ввод, duplicate в batch, повторный отдельный запуск,
unsupported demo profile, сумма, BTC, bps, отсутствующая конфигурация, UUID,
несовпадающий job_id, неправильный result, сетевой сбой и поздний ответ.
Запуск после распаковки: node run_checks.mjs
Это офлайн проверки функций, а не полноценный n8n runtime или backend E2E.
Архитектура и критерии production изложены в отдельном ТЗ UzWealth AI.
