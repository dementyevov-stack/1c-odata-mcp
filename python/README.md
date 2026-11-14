
## Пакетный исполнитель поступлений и СФ

`batch_posting.BatchExecutor` — один код для партий A/B одного профиля
`fresh-purchase-invoice-v1`. Вход: JSON-план с `schema_version`, профилем,
партией и парами; SHA-256 канонического плана проверяется при чтении
квитанций. Выход: каталог `batch/<id>/` с JSON-квитанциями и `summary.json` /
`summary.md`.

```python
from batch_posting import BatchExecutor

executor = BatchExecutor(transport, "receipts")
summary = executor.run(plan_json, batch_id="A")
# plan_json: {"schema_version":1,"profile":"fresh-purchase-invoice-v1","batches":[...]}
```

`transport` (fake HTTP в тестах) предоставляет `validate_context`,
`find_by_source`, `get_document`, `post(kind, ref)`, `list_register`,
`vat_book`, `journal` и может бросить `ResponseLost`/`Register404`.

- `ResponseLost` после Post: карточка/регистратор перечитываются по тому же
  Ref_Key; подтверждённый результат фиксируется в квитанции, второй POST не
  выполняется. Не найденный после потери ответа документ — `unknown`.
- Подтверждённая квитанция при повторе — `skipped`, без POST; `unknown` и
  `conflict` также не повторяют запись.
- Сверяются статус, неизменность карточки/таблиц (кроме технических полей),
  набор регистра по `Recorder`+`Recorder_Type`, адресные строки, НДС, книга
  покупок и журнал. `404` набора допустим только для явно пустого ожидаемого
  набора при подтверждённом posted-документе и пустых `_RecordType`/движениях.
- `expected.deduction` проверяется отдельно у каждого СФ; `advances` — точные
  ссылки и время.
