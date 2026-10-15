# Annotation APIs

This domain covers hyperlinks, legacy comments, and threaded comments.

Use this document when the task says: hyperlink, link, comment, note, threaded comment, author, comment text, or cell annotation.

## Skill Summary

Annotation APIs are async live-object APIs. Collections are exposed as worksheet properties such as `sheet.hyperlinks`, not methods. Item handles work in both inline and worker mode.

## Hyperlinks

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.hyperlinks.items()` | List hyperlinks as live handles. | none | `Promise<ExcelHyperlink[]>` |
| `sheet.hyperlinks.getItem(ref)` | Get hyperlink handle by cell/range ref. | `string` | `ExcelHyperlink` |
| `sheet.hyperlinks.getItemOrNullObject(ref)` | Get hyperlink handle by ref, or null. | `string` | `Promise<ExcelHyperlink \| null>` |
| `sheet.hyperlinks.add(model)` | Add a hyperlink from model data. | `ExcelHyperlinkModel` | `Promise<ExcelHyperlink>` |
| `sheet.hyperlinks.delete(ref)` | Delete hyperlink when present. | `string` | `Promise<void>` |
| `sheet.hyperlinks.toModel()` | Return hyperlink snapshots. | none | `Promise<ExcelHyperlinkModel[]>` |

`ExcelHyperlink`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `hyperlink.ref()` | Read anchor range. | none | `Promise<string>` |
| `hyperlink.setRef(ref)` | Set anchor range. | `string` | `Promise<void>` |
| `hyperlink.target()` | Read external target. | none | `Promise<string \| undefined>` |
| `hyperlink.targetMode()` | Read target mode. | none | `Promise<string \| undefined>` |
| `hyperlink.display()` | Read display text. | none | `Promise<string \| undefined>` |
| `hyperlink.setDisplay(display)` | Set or clear display text. | `string \| undefined` | `Promise<void>` |
| `hyperlink.setLocation(location)` | Set or clear internal location. | `string \| undefined` | `Promise<void>` |
| `hyperlink.toModel()` | Return a snapshot. | none | `Promise<ExcelHyperlinkModel>` |
| `hyperlink.delete()` | Remove this hyperlink. | none | `Promise<void>` |

## Legacy Comments

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.comments.items()` | List comments as live handles. | none | `Promise<ExcelComment[]>` |
| `sheet.comments.getItem(ref)` | Get comment handle by cell ref. | `string` | `ExcelComment` |
| `sheet.comments.getItemOrNullObject(ref)` | Get comment handle by ref, or null. | `string` | `Promise<ExcelComment \| null>` |
| `sheet.comments.add(model)` | Add a comment from model data. | `ExcelCommentModel` | `Promise<ExcelComment>` |
| `sheet.comments.delete(ref)` | Delete comment when present. | `string` | `Promise<void>` |
| `sheet.comments.toModel()` | Return comment snapshots. | none | `Promise<ExcelCommentModel[]>` |

`ExcelComment`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `comment.ref()` | Read cell ref. | none | `Promise<string>` |
| `comment.setRef(ref)` | Set cell ref. | `string` | `Promise<void>` |
| `comment.text()` | Read plain text. | none | `Promise<string>` |
| `comment.setText(text)` | Set plain text. | `string` | `Promise<void>` |
| `comment.toModel()` | Return a snapshot. | none | `Promise<ExcelCommentModel>` |
| `comment.delete()` | Remove this comment. | none | `Promise<void>` |

## Threaded Comments

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `sheet.threadedComments.items()` | List threaded comments as live handles. | none | `Promise<ExcelThreadedComment[]>` |
| `sheet.threadedComments.getItem(ref)` | Get threaded-comment handle by cell ref. | `string` | `ExcelThreadedComment` |
| `sheet.threadedComments.getItemOrNullObject(ref)` | Get threaded-comment handle by cell ref, or null. | `string` | `Promise<ExcelThreadedComment \| null>` |
| `sheet.threadedComments.add(model)` | Add a threaded comment from model data. | `ExcelThreadedCommentModel` | `Promise<ExcelThreadedComment>` |
| `sheet.threadedComments.delete(ref)` | Delete threaded comment when present. | `string` | `Promise<void>` |
| `sheet.threadedComments.toModel()` | Return threaded-comment snapshots. | none | `Promise<ExcelThreadedCommentModel[]>` |

`ExcelThreadedComment`:

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `threadedComment.ref()` | Read cell ref. | none | `Promise<string>` |
| `threadedComment.setRef(ref)` | Set cell ref. | `string` | `Promise<void>` |
| `threadedComment.text()` | Read plain text. | none | `Promise<string>` |
| `threadedComment.setText(text)` | Set plain text. | `string` | `Promise<void>` |
| `threadedComment.toModel()` | Return a snapshot. | none | `Promise<ExcelThreadedCommentModel>` |
| `threadedComment.delete()` | Remove this threaded comment. | none | `Promise<void>` |

## CodeAgent Rules

1. Use comments for cell notes and annotations; use threaded comments only when the workbook already has threaded-comment data or the task explicitly asks for it.
2. Use hyperlinks for link metadata; use `cell.setValue(...)` separately when display cell text must change.
3. If `setRef(...)` changes an item key, continue using the same handle or reacquire it by the new ref.
4. Deleting a row or column may remove comments/hyperlinks whose anchors are deleted.
