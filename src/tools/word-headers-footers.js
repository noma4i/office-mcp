import { validateString, validateInteger, validateNumber } from '../lib/validators.js';
import { runAppleScript } from '../lib/applescript/executor.js';
import { toAppleScriptString, escapeAppleScriptString } from '../lib/applescript/helpers.js';
import { wrapWordScript } from '../lib/applescript/script-wrappers.js';

function headerFooterIndex(type) {
  switch (type) {
    case 'first_page':
      return 'header footer first page';
    case 'even_pages':
      return 'header footer even pages';
    default:
      return 'header footer primary';
  }
}

export const headerFooterTools = [
  {
    name: 'word_get_header_text',
    description: 'Get header text content from a section in the active Word document',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        section: { type: 'integer', description: 'Section number (1-based, default: 1)', default: 1 },
        type: {
          type: 'string',
          description: 'Header type: "primary" (default), "first_page", or "even_pages"',
          enum: ['primary', 'first_page', 'even_pages'],
          default: 'primary'
        }
      }
    },
    async handler(args) {
      const section = validateInteger(args.section, 'section', 1) || 1;
      const hfIndex = headerFooterIndex(args.type);
      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${section} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set refHeader to get header of section ${section} of d index ${hfIndex}
  set headerText to content of text object of refHeader
on error
  return "Header not available for this section/type"
end try
return headerText
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_set_header_text',
    description: 'Set header text content in a section of the active Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Text to set as header content' },
        section: { type: 'integer', description: 'Section number (1-based, default: 1)', default: 1 },
        type: {
          type: 'string',
          description: 'Header type: "primary" (default), "first_page", or "even_pages"',
          enum: ['primary', 'first_page', 'even_pages'],
          default: 'primary'
        }
      },
      required: ['text']
    },
    async handler(args) {
      const text = validateString(args.text, 'text', true);
      const section = validateInteger(args.section, 'section', 1) || 1;
      const hfIndex = headerFooterIndex(args.type);
      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${section} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set refHeader to get header of section ${section} of d index ${hfIndex}
  set content of text object of refHeader to ${toAppleScriptString(text)}
on error
  return "Header not available for this section/type"
end try
return "Header text set for section ${section}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_get_footer_text',
    description: 'Get footer text content from a section in the active Word document',
    annotations: { readOnlyHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        section: { type: 'integer', description: 'Section number (1-based, default: 1)', default: 1 },
        type: {
          type: 'string',
          description: 'Footer type: "primary" (default), "first_page", or "even_pages"',
          enum: ['primary', 'first_page', 'even_pages'],
          default: 'primary'
        }
      }
    },
    async handler(args) {
      const section = validateInteger(args.section, 'section', 1) || 1;
      const hfIndex = headerFooterIndex(args.type);
      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${section} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set refFooter to get footer of section ${section} of d index ${hfIndex}
  set footerText to content of text object of refFooter
on error
  return "Footer not available for this section/type"
end try
return footerText
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_set_footer_text',
    description: 'Set footer text content in a section of the active Word document',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Text to set as footer content' },
        section: { type: 'integer', description: 'Section number (1-based, default: 1)', default: 1 },
        type: {
          type: 'string',
          description: 'Footer type: "primary" (default), "first_page", or "even_pages"',
          enum: ['primary', 'first_page', 'even_pages'],
          default: 'primary'
        }
      },
      required: ['text']
    },
    async handler(args) {
      const text = validateString(args.text, 'text', true);
      const section = validateInteger(args.section, 'section', 1) || 1;
      const hfIndex = headerFooterIndex(args.type);
      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${section} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set refFooter to get footer of section ${section} of d index ${hfIndex}
  set content of text object of refFooter to ${toAppleScriptString(text)}
on error
  return "Footer not available for this section/type"
end try
return "Footer text set for section ${section}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_insert_header_image',
    description: 'Insert an image into the header of a Word document section',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Full path to the image file' },
        section: { type: 'integer', description: 'Section number (1-based, default: 1)', default: 1 },
        type: {
          type: 'string',
          description: 'Header type: "primary" (default), "first_page", or "even_pages"',
          enum: ['primary', 'first_page', 'even_pages'],
          default: 'primary'
        },
        width: { type: 'number', description: 'Optional width in points to resize the image' },
        height: { type: 'number', description: 'Optional height in points to resize the image' }
      },
      required: ['path']
    },
    async handler(args) {
      const path = validateString(args.path, 'path', true);
      const section = validateInteger(args.section, 'section', 1) || 1;
      const hfIndex = headerFooterIndex(args.type);
      const posixPath = path.startsWith('/') ? path : `/${path}`;
      const hfsPath = posixPath.replace(/\//g, ':').replace(/^:/, '');
      let resizeCommands = '';
      if (args.width !== undefined) resizeCommands += `\n  set width of shp to ${validateNumber(args.width, 'width', 1, 10000)}`;
      if (args.height !== undefined) resizeCommands += `\n  set height of shp to ${validateNumber(args.height, 'height', 1, 10000)}`;

      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${section} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set refHeader to get header of section ${section} of d index ${hfIndex}
  set hdrRange to text object of refHeader
on error
  return "Header not available for this section/type"
end try
set shp to make new inline picture at end of hdrRange with properties {file name:"${escapeAppleScriptString(hfsPath)}", save with document:true}${resizeCommands}
return "Image inserted into header of section ${section}"
`);
      return await runAppleScript(script);
    }
  },
  {
    name: 'word_insert_footer_image',
    description: 'Insert an image into the footer of a Word document section',
    annotations: { destructiveHint: true },
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Full path to the image file' },
        section: { type: 'integer', description: 'Section number (1-based, default: 1)', default: 1 },
        type: {
          type: 'string',
          description: 'Footer type: "primary" (default), "first_page", or "even_pages"',
          enum: ['primary', 'first_page', 'even_pages'],
          default: 'primary'
        },
        width: { type: 'number', description: 'Optional width in points to resize the image' },
        height: { type: 'number', description: 'Optional height in points to resize the image' }
      },
      required: ['path']
    },
    async handler(args) {
      const path = validateString(args.path, 'path', true);
      const section = validateInteger(args.section, 'section', 1) || 1;
      const hfIndex = headerFooterIndex(args.type);
      const posixPath = path.startsWith('/') ? path : `/${path}`;
      const hfsPath = posixPath.replace(/\//g, ':').replace(/^:/, '');
      let resizeCommands = '';
      if (args.width !== undefined) resizeCommands += `\n  set width of shp to ${validateNumber(args.width, 'width', 1, 10000)}`;
      if (args.height !== undefined) resizeCommands += `\n  set height of shp to ${validateNumber(args.height, 'height', 1, 10000)}`;

      const script = wrapWordScript(`
set d to active document
set secCount to count of sections of d
if ${section} > secCount then
  return "Section index out of range. Document has " & secCount & " sections."
end if
try
  set refFooter to get footer of section ${section} of d index ${hfIndex}
  set ftrRange to text object of refFooter
on error
  return "Footer not available for this section/type"
end try
set shp to make new inline picture at end of ftrRange with properties {file name:"${escapeAppleScriptString(hfsPath)}", save with document:true}${resizeCommands}
return "Image inserted into footer of section ${section}"
`);
      return await runAppleScript(script);
    }
  }
];

