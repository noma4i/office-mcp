import { validateString } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const bookmarkTools = [
  {
    name: 'list_bookmarks',
    description: 'List all bookmarks in the active document',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {}
    },
    async handler() {
      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
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
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'create_bookmark',
    description: 'Create a bookmark at the current selection',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Bookmark name'
        }
      },
      required: ['name']
    },
    async handler(args) {
      const name = validateString(args.name, 'name', true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          make new bookmark at d with properties {name:${JSON.stringify(name)}, |bookmark range|:selection}
          return "Bookmark created: " & ${JSON.stringify(name)}
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'goto_bookmark',
    description: 'Jump to a bookmark by name',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Bookmark name'
        }
      },
      required: ['name']
    },
    async handler(args) {
      const name = validateString(args.name, 'name', true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set b to bookmark ${JSON.stringify(name)} of d
          select (text object of b)
          return "Jumped to bookmark: " & ${JSON.stringify(name)}
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'delete_bookmark',
    description: 'Delete a bookmark by name',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Bookmark name'
        }
      },
      required: ['name']
    },
    async handler(args) {
      const name = validateString(args.name, 'name', true);

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          delete bookmark ${JSON.stringify(name)} of d
          return "Bookmark deleted: " & ${JSON.stringify(name)}
        end tell
      `;

      return await runAppleScript(script);
    }
  }
];
