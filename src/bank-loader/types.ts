import { z } from "zod";

/**
 * Первый участок загрузчика банковских документов БП: строка выписки (источник) и целевые
 * реквизиты оператора. Реквизиты задаются локально и не попадают в Git.
 */

const inn = z
  .string()
  .trim()
  .regex(/^\d{10}$|^\d{12}$/, "ИНН должен содержать 10 или 12 цифр");
const kpp = z
  .string()
  .trim()
  .regex(/^\d{9}$/, "КПП должен содержать 9 цифр");
const accountNumber = z
  .string()
  .trim()
  .regex(/^\d{20}$/, "Номер счёта должен содержать 20 цифр");

export const TargetRequisitesSchema = z.object({
  organizationInn: inn,
  organizationKpp: kpp.optional(),
  bankAccountNumber: accountNumber,
  /** БИК банка организации: нужен только если счёт организации придётся создать. */
  bankBik: z
    .string()
    .trim()
    .regex(/^\d{9}$/, "БИК должен содержать 9 цифр")
    .optional(),
});
export type TargetRequisites = z.infer<typeof TargetRequisitesSchema>;

export const SourceDocumentSchema = z.object({
  /** Стабильный ключ источника: по нему загрузчик связывает строку выписки с Ref_Key документа. */
  sourceKey: z.string().trim().min(1).max(200),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/, "Дата документа — YYYY-MM-DD"),
  amount: z.number().positive("Сумма должна быть больше нуля"),
  payerInn: inn,
  payerKpp: kpp.optional(),
  payerName: z.string().trim().max(300).optional(),
  /** Счёт плательщика из выписки: переносится в документ, если он опубликован в схеме. */
  payerAccountNumber: accountNumber.optional(),
  purpose: z.string().trim().max(500).optional(),
  /** Номер и дата входящего банковского документа: переносятся явно, не угадываются. */
  documentNumber: z.string().trim().min(1).max(100).optional(),
  documentDate: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/, "Дата входящего документа — YYYY-MM-DD")
    .optional(),
  /** Явная историческая ставка и сумма НДС из выписки; обе или ни одной. */
  vatRate: z.string().trim().min(1).max(50).optional(),
  vatAmount: z.number().min(0).optional(),
  contractNumber: z.string().trim().min(1).max(100).optional(),
  contractRef: z.string().uuid("contractRef должен быть UUID").optional(),
});
export type SourceDocument = z.infer<typeof SourceDocumentSchema>;

export const SourcePackageSchema = z.union([
  z.object({ documents: z.array(SourceDocumentSchema).min(1) }),
  z.array(SourceDocumentSchema).min(1),
]);
export type SourcePackage = z.infer<typeof SourcePackageSchema>;

export const LoaderConfigSchema = z
  .object({
    odata: z.object({
      baseUrl: z.string().url("odata.baseUrl должен быть URL"),
      username: z.string().min(1).optional(),
      password: z.string().min(1).optional(),
      usernameEnv: z.string().min(1).optional(),
      passwordEnv: z.string().min(1).optional(),
      /** Журнал операций записи MCP: должен переживать перезапуск загрузчика. */
      writeJournalDir: z.string().min(1).optional(),
    }),
    target: TargetRequisitesSchema,
    /** Журнал загрузчика: сопоставление sourceKey → Ref_Key и результаты. Не коммитить. */
    journalDir: z.string().min(1).default("~/.1c-odata-mcp/bank-loader"),
  })
  .strict();
export type LoaderConfig = z.infer<typeof LoaderConfigSchema>;
