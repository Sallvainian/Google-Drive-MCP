# Google Docs MCP Server

FastMCP server with 58 tools for Google Docs, Sheets, Slides, and Drive.

## Tool Categories

| Category | Count | Examples |
|----------|-------|----------|
| Docs | 5 | `readGoogleDoc`, `appendToGoogleDoc`, `insertText`, `deleteRange`, `listDocumentTabs` |
| Formatting | 3 | `applyTextStyle`, `applyParagraphStyle`, `formatMatchingText` |
| Structure | 7 | `insertTable`, `insertPageBreak`, `insertImageFromUrl`, `insertLocalImage`, `editTableCell`*, `findElement`*, `fixListFormatting`* |
| Comments | 6 | `listComments`, `getComment`, `addComment`, `replyToComment`, `resolveComment`, `deleteComment` |
| Sheets | 8 | `readSpreadsheet`, `writeSpreadsheet`, `appendSpreadsheetRows`, `clearSpreadsheetRange`, `createSpreadsheet`, `listGoogleSheets` |
| Drive | 13 | `listGoogleDocs`, `searchGoogleDocs`, `getDocumentInfo`, `createFolder`, `moveFile`, `copyFile`, `createDocument` |
| Slides | 16 | `getPresentation`, `listSlides`, `getSlide`, `mapSlide`, `createPresentation`, `addSlide`, `duplicateSlide`, `addTextBox`, `addShape`, `addImage`, `addTable`, `deleteSlide`, `deleteElement`, `updateSpeakerNotes`, `moveSlide`, `insertTextInElement` |

*Not fully implemented

## Known Limitations

- **Comment anchoring:** Programmatically created comments appear in "All Comments" but aren't visibly anchored to text in the UI
- **Resolved status:** May not persist in Google Docs UI (Drive API limitation)
- **editTableCell:** Not implemented (complex cell index calculation)
- **fixListFormatting:** Experimental, may not work reliably

## Parameter Patterns

- **Document ID:** Extract from URL: `docs.google.com/document/d/DOCUMENT_ID/edit`
- **Presentation ID:** Extract from URL: `docs.google.com/presentation/d/PRESENTATION_ID/edit`
- **Text targeting:** Use `textToFind` + `matchInstance` OR `startIndex`/`endIndex`
- **Colors:** Hex format `#RRGGBB` or `#RGB`
- **Alignment:** `START`, `END`, `CENTER`, `JUSTIFIED` (not LEFT/RIGHT)
- **Indices:** 1-based for Docs, 0-based for Slides positions
- **Tabs:** Optional `tabId` parameter (defaults to first tab)
- **Slides positioning:** All positions and sizes in points (1 inch = 72 points)
- **Slides workflow:** Use `mapSlide` BEFORE adding elements to see dimensions and available space
- **Text box styling:** `addTextBox` supports `fontSize` (points), `fontFamily`, `bold`, `italic`, `textColor` (hex)
- **getSlide modes:** Single slide by `pageObjectId`, OR range by `startIndex`/`endIndex` (0-based, exclusive end)

## Source Files (for implementation details)

| File | Contains |
|------|----------|
| `src/types.ts` | Zod schemas, hex color validation, style parameter definitions |
| `src/googleDocsApiHelpers.ts` | `findTextRange`, `executeBatchUpdate`, style request builders |
| `src/googleSheetsApiHelpers.ts` | A1 notation parsing, range operations |
| `src/googleSlidesApiHelpers.ts` | EMU conversion, shape/table builders, batch update helpers |
| `src/server.ts` | All 58 tool definitions with full parameter schemas |

## See Also

- `README.md` - Setup instructions and usage examples
- `SAMPLE_TASKS.md` - 15 example workflows
