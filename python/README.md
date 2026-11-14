# Python-проверка справочников

Минимальный независимый патч для Python-исполнителя: без HTTP-клиента, без
рабочих данных и без второго механизма записи.

## Интеграция цикла

Рабочий цикл — `execute_catalog_cycle(adapter, plan)` в
`catalog_verifier.py`. Он не знает HTTP: исполнитель передаёт адаптер с четырьмя
операциями, а fake-адаптер используется в тестах.

```python
from catalog_verifier import CatalogAdapter, CatalogCyclePlan, execute_catalog_cycle

adapter = CatalogAdapter(
    find_exact=executor.find_exact,       # точный поиск по source_key
    post=executor.post_catalog_row,       # POST
    get=executor.get_catalog_row,         # GET по Ref_Key
    patch=executor.patch_catalog_row,     # точечный PATCH
)
plan = CatalogCyclePlan(
    source_key="SRC-1",
    create_payload={"Code": "SRC-1", "Description": "synthetic", "Optional": None},
    expected_existing={"Code": "SRC-1", "Description": "synthetic"},
    expected_created={"Code": "SRC-1", "Description": "synthetic", "Optional": None},
    patch_changes={"DisplayOrder": 2},
    expected_patched={"DisplayOrder": 2},
    ignore_fields=("DataVersion",),
)
result = execute_catalog_cycle(adapter, plan)
assert result.status in {"existing", "created", "patched", "unknown", "conflict"}
```

- Адаптер должен бросать `ResponseLostError` при обрыве POST/PATCH и
  `ReadUnavailableError` при недоступном GET. Цикл возвращает `unknown`; только
  повторный GET по тому же `Ref_Key` может закрыть неизвестность, POST не
  повторяется.
- `has_fields` и `changed_fields` — строительные блоки цикла: наличие каждого
  ключа, точный тип/значение, пропажа null-поля и постороннее изменение.
- Нормализация передаётся только для конкретного известного поля, например
  `has_fields(row, {"DataVersion": 123}, {"DataVersion": str})` для Edm.Int64
  в виде строки. Глобальной нормализации нет.
- При точном найденном объекте сначала выполняется GET по найденному
  `Ref_Key`, проверяется идентичность и все конечные постусловия; пустой/чужой
  Ref_Key, недоступный GET или старое значение — `conflict`/`unknown`, без POST.
  Восстановление после потерянного PATCH использует уже известный `Ref_Key`, без
  нового поиска. Неоднозначный или неполный поиск — `conflict`, тоже без POST.

## Тесты

```bash
python3 -m unittest discover -s python/tests -v
```
