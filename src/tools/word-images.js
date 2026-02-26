import { validateString, validateNumber, validateInteger, validateBoolean } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';

export const imageTools = [
  {
    name: 'word_insert_image',
    description: 'Insert an image into the Word document at the current cursor position using macOS clipboard',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        path: {
          type: 'string',
          description: 'Full path to the image file (PNG, JPEG, TIFF, etc.)'
        },
        width: {
          type: 'number',
          description: 'Optional width in points to resize the image after insertion'
        },
        height: {
          type: 'number',
          description: 'Optional height in points to resize the image after insertion'
        }
      },
      required: ['path']
    },
    async handler(args) {
      const path = validateString(args.path, 'path', true);
      const width = args.width !== undefined ? validateNumber(args.width, 'width', 1, 10000) : undefined;
      const height = args.height !== undefined ? validateNumber(args.height, 'height', 1, 10000) : undefined;

      let resizeBlock = '';
      if (width !== undefined || height !== undefined) {
        resizeBlock = `
          set shp to inline shape (count of inline shapes of d) of d
          set lock aspect ratio of shp to ${width !== undefined && height !== undefined ? 'false' : 'true'}`;
        if (width !== undefined) {
          resizeBlock += `\n          set width of shp to ${width}`;
        }
        if (height !== undefined) {
          resizeBlock += `\n          set height of shp to ${height}`;
        }
      }

      const script = `
        set imgFile to POSIX file ${JSON.stringify(path)}
        try
          set the clipboard to (read imgFile as «class PNGf»)
        on error
          try
            set the clipboard to (read imgFile as TIFF picture)
          on error
            try
              set the clipboard to (read imgFile as JPEG picture)
            on error errMsg
              return "Error reading image: " & errMsg
            end try
          end try
        end try
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set shapesBefore to count of inline shapes of d
          paste object selection
          set shapesAfter to count of inline shapes of d
          if shapesAfter = shapesBefore then
            return "Image may not have been inserted. Try a different image format."
          end if${resizeBlock}
          return "Image inserted successfully. Total inline shapes: " & shapesAfter
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_list_inline_shapes',
    description: 'List all inline shapes (images, objects) in the active Word document with their dimensions',
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
          set shapeCount to count of inline shapes of d
          if shapeCount = 0 then
            return "No inline shapes found"
          end if
          set shapeList to ""
          repeat with i from 1 to shapeCount
            try
              set shp to inline shape i of d
              set t to inline shape type of shp
              set w to width of shp
              set h to height of shp
              set alt to ""
              try
                set alt to alternative text of shp
              end try
              set shapeList to shapeList & i & ". type=" & (t as text) & ", width=" & (w as text) & "pt, height=" & (h as text) & "pt"
              if alt is not "" then
                set shapeList to shapeList & ", alt=" & alt
              end if
              set shapeList to shapeList & linefeed
            on error
              set shapeList to shapeList & i & ". (not accessible)" & linefeed
            end try
          end repeat
          return shapeList
        end tell
      `;

      return await runAppleScript(script);
    }
  },

  {
    name: 'word_resize_inline_shape',
    description: 'Resize an inline shape (image) by index in the active Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        index: {
          type: 'integer',
          description: 'Inline shape index (1-based)'
        },
        width: {
          type: 'number',
          description: 'New width in points'
        },
        height: {
          type: 'number',
          description: 'New height in points'
        },
        lockAspectRatio: {
          type: 'boolean',
          description: 'Lock aspect ratio when resizing (default: true)',
          default: true
        }
      },
      required: ['index']
    },
    async handler(args) {
      const index = validateInteger(args.index, 'index', 1, 10000);
      const width = args.width !== undefined ? validateNumber(args.width, 'width', 1, 10000) : undefined;
      const height = args.height !== undefined ? validateNumber(args.height, 'height', 1, 10000) : undefined;
      const lockAspectRatio = args.lockAspectRatio !== undefined ? validateBoolean(args.lockAspectRatio, 'lockAspectRatio') : true;

      if (width === undefined && height === undefined) {
        throw new Error('At least one of width or height is required');
      }

      let resizeCommands = [];
      resizeCommands.push(`set lock aspect ratio of shp to ${lockAspectRatio}`);
      if (width !== undefined) {
        resizeCommands.push(`set width of shp to ${width}`);
      }
      if (height !== undefined) {
        resizeCommands.push(`set height of shp to ${height}`);
      }

      const script = `
        tell application "Microsoft Word"
          if (count of documents) = 0 then
            return "No document is open"
          end if
          set d to active document
          set shapeCount to count of inline shapes of d
          if ${index} > shapeCount then
            return "Shape index out of range. Document has " & shapeCount & " inline shapes."
          end if
          try
            set shp to inline shape ${index} of d
            ${resizeCommands.join('\n            ')}
          on error errMsg
            return "Error resizing shape ${index}: " & errMsg
          end try
          return "Shape ${index} resized. Width: " & (width of shp as text) & "pt, Height: " & (height of shp as text) & "pt"
        end tell
      `;

      return await runAppleScript(script);
    }
  }
];
