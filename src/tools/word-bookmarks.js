import { validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { quoteAppleScriptString } from '../lib/applescript/helpers.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

export const bookmarkTools = [
  {
    name: 'word_list_bookmarks',
    description: 'List all bookmarks in the active Word document',
    annotations: { readOnlyHint: true },
    inputSchema: { type: 'object', properties: {} },
    async handler() {
      const script = wrapWordScript(`
set d to active document
set bookmarkCount to count of bookmarks of d
if bookmarkCount = 0 then
  return "No bookmarks found"
end if
set bookmarkList to ""
repeat with i from 1 to bookmarkCount
  set b to bookmark i of d
  set bookmarkList to bookmarkList & i & ". " & (name of b) & linefeed
end repeat
return bookmarkList
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_create_bookmark',
    description: 'Create a bookmark at the current selection in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: { name: { type: 'string', description: 'Bookmark name' } },
      required: ['name']
    },
    async handler(args) {
      const name = validateString(args.name, 'name', true);
      const script = wrapWordScript(`
set d to active document
try
  make new bookmark at d with properties {name:${quoteAppleScriptString(name)}, |bookmark range|:selection}
on error errMsg
  return "Error creating bookmark: " & errMsg
end try
return "Bookmark created: " & ${quoteAppleScriptString(name)}
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_goto_bookmark',
    description: 'Jump to a bookmark by name in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: { name: { type: 'string', description: 'Bookmark name' } },
      required: ['name']
    },
    async handler(args) {
      const name = validateString(args.name, 'name', true);
      const script = wrapWordScript(`
set d to active document
try
  set b to bookmark ${quoteAppleScriptString(name)} of d
  select (text object of b)
on error
  return "Bookmark not found: " & ${quoteAppleScriptString(name)}
end try
return "Jumped to bookmark: " & ${quoteAppleScriptString(name)}
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_delete_bookmark',
    description: 'Delete a bookmark by name in Word',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: { name: { type: 'string', description: 'Bookmark name' } },
      required: ['name']
    },
    async handler(args) {
      const name = validateString(args.name, 'name', true);
      const script = wrapWordScript(`
set d to active document
try
  delete bookmark ${quoteAppleScriptString(name)} of d
on error
  return "Bookmark not found: " & ${quoteAppleScriptString(name)}
end try
return "Bookmark deleted: " & ${quoteAppleScriptString(name)}
`);
      return await runAppleScript(script);
    }
  }
];

