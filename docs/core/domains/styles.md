# Style APIs

This domain covers reading style definitions from `styles.xml`.

Use this document when the task says: style, heading style, paragraph style, character style, table style, numbering style, basedOn, outline level, or style id.

## Skill Summary

Style APIs are read-only. They help code decide which paragraphs or tables match a style. Applying styles is not exposed as a direct style-object write API in the current live-object surface.

## Entry Points

| API | Purpose | Params | Returns |
| --- | --- | --- | --- |
| `doc.styles.list(opts?)` | List styles. | `ListStylesOptions` | `Promise<StyleInfo[]>` |
| `doc.styles.get(styleId)` | Get one style by id. | `styleId: string` | `Promise<StyleInfo \| null>` |

## ListStylesOptions

| Field | Type | Default | Description |
| --- | --- | --- | --- |
| `type` | `'paragraph' \| 'character' \| 'table' \| 'numbering'` | none | Filter styles by OOXML style type. |

## StyleInfo

| Field | Type | Description |
| --- | --- | --- |
| `styleId` | `string` | Style id used by paragraph/table references. Case-sensitive. |
| `name` | `string` | Display/name value. If missing, falls back to `styleId`. |
| `type` | `'paragraph' \| 'character' \| 'table' \| 'numbering'` | OOXML style type. Unknown values fall back to `paragraph`. |
| `basedOn` | `string?` | Parent style id. Core does not validate inheritance cycles. |
| `outlineLevel` | `number?` | Heading level marker from the style definition. `0` maps to heading 1. |

## Examples

Find heading-like paragraph styles:

```ts
const styles = await doc.styles.list({ type: 'paragraph' });
const headingStyles = styles.filter((style) => style.outlineLevel !== undefined);
```

List paragraphs using a style:

```ts
const style = await doc.styles.get('Heading1');
if (style) {
  const result = await doc.paragraphs.list({
    filter: { styleId: style.styleId },
    pagination: { mode: 'all' },
  });
}
```

## CodeAgent Rules

1. Use styles to identify document structure; do not invent `style.apply()` or `doc.styles.set`.
2. `styleId` is case-sensitive.
3. `basedOn` is a raw reference; generated code must avoid recursive traversal without cycle protection.
4. Style values do not include fully resolved inherited formatting.
