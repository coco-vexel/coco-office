# Formatting Rule APIs

This domain covers data validations and conditional formats.

Use this document when the task says: data validation, dropdown, allowed values, validation prompt, conditional formatting, color scale, cell rule, or format rule.

## Skill Summary

Formatting-rule APIs are async live-object APIs. Collections are exposed as worksheet properties such as `sheet.dataValidations` and `sheet.conditionalFormats`.

## Data Validations

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.dataValidations.items()` | List data validations as live handles. | none | `Promise<ExcelDataValidation[]>` |
| `sheet.dataValidations.getItem(rangeAddress)` | Get validation handle by target range. | `string` | `ExcelDataValidation` |
| `sheet.dataValidations.getItemOrNullObject(rangeAddress)` | Get validation handle by target range, or null. | `string` | `Promise<ExcelDataValidation \| null>` |
| `sheet.dataValidations.add(model)` | Add a validation from model data. | `ExcelDataValidationModel` | `Promise<ExcelDataValidation>` |
| `sheet.dataValidations.delete(rangeAddress)` | Delete validation by target range when present. | `string` | `Promise<void>` |
| `sheet.dataValidations.toModel()` | Return validation snapshots. | none | `Promise<ExcelDataValidationModel[]>` |

`ExcelDataValidation`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `validation.rangeAddress()` | Read target range list. | none | `Promise<string \| undefined>` |
| `validation.setRange(address)` | Set target range list. | `string` | `Promise<void>` |
| `validation.type()` | Read validation type. | none | `Promise<string \| undefined>` |
| `validation.setType(type)` | Set or clear validation type. | `string \| undefined` | `Promise<void>` |
| `validation.operator()` | Read validation operator. | none | `Promise<string \| undefined>` |
| `validation.setOperator(operator)` | Set or clear validation operator. | `string \| undefined` | `Promise<void>` |
| `validation.allowBlank()` | Read allow-blank flag. | none | `Promise<boolean \| undefined>` |
| `validation.setAllowBlank(value)` | Set or clear allow-blank flag. | `boolean \| undefined` | `Promise<void>` |
| `validation.showDropDown()` | Read dropdown visibility flag. | none | `Promise<boolean \| undefined>` |
| `validation.setShowDropDown(value)` | Set or clear dropdown visibility flag. | `boolean \| undefined` | `Promise<void>` |
| `validation.showInputMessage()` | Read input-message flag. | none | `Promise<boolean \| undefined>` |
| `validation.setShowInputMessage(value)` | Set or clear input-message flag. | `boolean \| undefined` | `Promise<void>` |
| `validation.showErrorMessage()` | Read error-message flag. | none | `Promise<boolean \| undefined>` |
| `validation.setShowErrorMessage(value)` | Set or clear error-message flag. | `boolean \| undefined` | `Promise<void>` |
| `validation.errorStyle()` | Read error style. | none | `Promise<string \| undefined>` |
| `validation.setErrorStyle(style)` | Set or clear error style. | `string \| undefined` | `Promise<void>` |
| `validation.errorTitle()` | Read error title. | none | `Promise<string \| undefined>` |
| `validation.error()` | Read error message. | none | `Promise<string \| undefined>` |
| `validation.setError(title, message)` | Set or clear error title/message. | `string \| undefined`, `string \| undefined` | `Promise<void>` |
| `validation.promptTitle()` | Read prompt title. | none | `Promise<string \| undefined>` |
| `validation.prompt()` | Read prompt message. | none | `Promise<string \| undefined>` |
| `validation.setPrompt(title, message)` | Set or clear prompt title/message. | `string \| undefined`, `string \| undefined` | `Promise<void>` |
| `validation.formula1()` | Read first validation formula. | none | `Promise<string \| undefined>` |
| `validation.setFormula1(formula)` | Set or clear first formula. | `string \| undefined` | `Promise<void>` |
| `validation.formula2()` | Read second validation formula. | none | `Promise<string \| undefined>` |
| `validation.setFormula2(formula)` | Set or clear second formula. | `string \| undefined` | `Promise<void>` |
| `validation.toModel()` | Return a snapshot. | none | `Promise<ExcelDataValidationModel>` |
| `validation.delete()` | Remove this validation. | none | `Promise<void>` |

## Conditional Formats

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.conditionalFormats.items()` | List conditional-format blocks. | none | `Promise<ExcelConditionalFormat[]>` |
| `sheet.conditionalFormats.getItem(rangeAddress)` | Get conditional-format block handle by target range. | `string` | `ExcelConditionalFormat` |
| `sheet.conditionalFormats.getItemOrNullObject(rangeAddress)` | Get conditional-format block handle by target range, or null. | `string` | `Promise<ExcelConditionalFormat \| null>` |
| `sheet.conditionalFormats.add(model)` | Add a conditional-format block. | `ExcelConditionalFormatModel` | `Promise<ExcelConditionalFormat>` |
| `sheet.conditionalFormats.delete(rangeAddress)` | Delete block by target range when present. | `string` | `Promise<void>` |
| `sheet.conditionalFormats.toModel()` | Return conditional-format snapshots. | none | `Promise<ExcelConditionalFormatModel[]>` |

`ExcelConditionalFormat`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `format.rangeAddress()` | Read target range list. | none | `Promise<string \| undefined>` |
| `format.setRange(address)` | Set target range list. | `string` | `Promise<void>` |
| `format.rules` | Conditional-format rule collection. | property | `ExcelConditionalFormatRuleCollection` |
| `format.toModel()` | Return a conditional-format snapshot. | none | `Promise<ExcelConditionalFormatModel>` |
| `format.delete()` | Remove this conditional-format block. | none | `Promise<void>` |

`ExcelConditionalFormatRuleCollection`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `format.rules.items()` | List rules. | none | `Promise<ExcelConditionalFormatRule[]>` |
| `format.rules.getItem(index)` | Get rule by zero-based index. | `number` | `ExcelConditionalFormatRule` |
| `format.rules.getItemOrNullObject(index)` | Get rule by zero-based index, or null. | `number` | `Promise<ExcelConditionalFormatRule \| null>` |
| `format.rules.add(model)` | Add a rule. | `ExcelConditionalFormatRuleModel` | `Promise<ExcelConditionalFormatRule>` |
| `format.rules.delete(index)` | Delete rule by zero-based index. | `number` | `Promise<void>` |
| `format.rules.toModel()` | Return rule snapshots. | none | `Promise<ExcelConditionalFormatRuleModel[]>` |

`ExcelConditionalFormatRule`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `rule.type()` | Read rule type. | none | `Promise<string \| undefined>` |
| `rule.setType(type)` | Set or clear rule type. | `string \| undefined` | `Promise<void>` |
| `rule.priority()` | Read priority. | none | `Promise<number \| undefined>` |
| `rule.setPriority(priority)` | Set or clear priority. | `number \| undefined` | `Promise<void>` |
| `rule.operator()` | Read operator. | none | `Promise<string \| undefined>` |
| `rule.setOperator(operator)` | Set or clear operator. | `string \| undefined` | `Promise<void>` |
| `rule.stopIfTrue()` | Read stop-if-true flag. | none | `Promise<boolean \| undefined>` |
| `rule.setStopIfTrue(value)` | Set or clear stop-if-true flag. | `boolean \| undefined` | `Promise<void>` |
| `rule.formulas()` | Read formulas. | none | `Promise<string[]>` |
| `rule.setFormula(index, formula)` | Replace one formula by zero-based index. | `number`, `string` | `Promise<void>` |
| `rule.addFormula(formula)` | Append a formula. | `string` | `Promise<void>` |
| `rule.deleteFormula(index)` | Delete one formula by zero-based index. | `number` | `Promise<void>` |
| `rule.toModel()` | Return a rule snapshot. | none | `Promise<ExcelConditionalFormatRuleModel>` |
| `rule.delete()` | Remove this rule. | none | `Promise<void>` |

## CodeAgent Rules

1. `rangeAddress()` maps to model `sqref`; it may contain multiple space-separated refs.
2. Use formulas without assuming a leading `=`.
3. If `setRange(...)` changes an item key, continue using the same handle or reacquire it by the new range.
4. Insert/delete row/column operations best-effort rewrite validation and conditional-format references.
5. Deleting validation or conditional-format metadata does not delete worksheet cells.
