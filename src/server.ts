import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { docs_v1, drive_v3, sheets_v4 } from "googleapis";

export const EXPORT_MIME_TYPES: Record<string, { mimeType: string; ext: string }> = {
  "application/vnd.google-apps.document": { mimeType: "text/markdown", ext: "md" },
  "application/vnd.google-apps.spreadsheet": { mimeType: "text/csv", ext: "csv" },
  "application/vnd.google-apps.presentation": { mimeType: "text/plain", ext: "txt" },
  "application/vnd.google-apps.drawing": { mimeType: "image/png", ext: "png" },
};

export interface DriveClients {
  drive: drive_v3.Drive;
  docs: docs_v1.Docs;
  sheets: sheets_v4.Sheets;
}

// ---------------------------------------------------------------------------
// Markdown → Google Docs format helpers
// ---------------------------------------------------------------------------

type NamedStyleType =
  | "HEADING_1" | "HEADING_2" | "HEADING_3"
  | "HEADING_4" | "HEADING_5" | "HEADING_6"
  | "NORMAL_TEXT";

interface TextRun { text: string; bold: boolean; italic: boolean }
interface ParsedParagraph { namedStyleType: NamedStyleType; plainText: string; runs: TextRun[] }

function parseInlineRuns(text: string): TextRun[] {
  if (!text) return [{ text: "", bold: false, italic: false }];
  const runs: TextRun[] = [];
  // Order matters: *** before ** before *
  const re = /\*\*\*([^*]+)\*\*\*|\*\*([^*]+)\*\*|\*([^*]+)\*|_([^_]+)_|([^*_]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[1] !== undefined) runs.push({ text: m[1], bold: true,  italic: true  });
    else if (m[2] !== undefined) runs.push({ text: m[2], bold: true,  italic: false });
    else if (m[3] !== undefined) runs.push({ text: m[3], bold: false, italic: true  });
    else if (m[4] !== undefined) runs.push({ text: m[4], bold: false, italic: true  });
    else if (m[5] !== undefined) runs.push({ text: m[5], bold: false, italic: false });
  }
  return runs.length > 0 ? runs : [{ text, bold: false, italic: false }];
}

function parseMarkdownParagraphs(markdown: string): ParsedParagraph[] {
  const lines = markdown.split("\n");
  while (lines.length > 0 && lines[lines.length - 1].trim() === "") lines.pop();
  if (lines.length === 0) return [{ namedStyleType: "NORMAL_TEXT", plainText: "", runs: [{ text: "", bold: false, italic: false }] }];

  return lines.map((line) => {
    const h = line.match(/^(#{1,6})\s+(.*)/);
    if (h) {
      const styleType = `HEADING_${h[1].length}` as NamedStyleType;
      return { namedStyleType: styleType, plainText: h[2], runs: parseInlineRuns(h[2]) };
    }
    return { namedStyleType: "NORMAL_TEXT", plainText: line, runs: parseInlineRuns(line) };
  });
}

function buildMarkdownRequests(paragraphs: ParsedParagraph[], bodyEndIndex: number): docs_v1.Schema$Request[] {
  const requests: docs_v1.Schema$Request[] = [];

  if (bodyEndIndex > 2) {
    requests.push({ deleteContentRange: { range: { startIndex: 1, endIndex: bodyEndIndex - 1 } } });
  }

  let idx = 1;
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const isLast = i === paragraphs.length - 1;
    const insertText = para.plainText + (isLast ? "" : "\n");

    if (insertText.length > 0) {
      requests.push({ insertText: { location: { index: idx }, text: insertText } });
    }

    // Style range always includes the trailing \n (inserted or the doc's existing final one)
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: idx, endIndex: idx + para.plainText.length + 1 },
        paragraphStyle: { namedStyleType: para.namedStyleType },
        fields: "namedStyleType",
      },
    });

    // Inline bold / italic
    let runIdx = idx;
    for (const run of para.runs) {
      if (run.text && (run.bold || run.italic)) {
        const fields = [run.bold && "bold", run.italic && "italic"].filter(Boolean).join(",");
        requests.push({
          updateTextStyle: {
            range: { startIndex: runIdx, endIndex: runIdx + run.text.length },
            textStyle: { ...(run.bold && { bold: true }), ...(run.italic && { italic: true }) },
            fields,
          },
        });
      }
      if (run.text) runIdx += run.text.length;
    }

    idx += insertText.length;
  }

  return requests;
}

export async function getFileContent(
  drive: drive_v3.Drive,
  fileId: string,
  mimeType: string
): Promise<{ content: string; mimeType: string }> {
  const exportFormat = EXPORT_MIME_TYPES[mimeType];
  if (exportFormat) {
    const res = await drive.files.export(
      { fileId, mimeType: exportFormat.mimeType },
      { responseType: "text" }
    );
    return { content: res.data as string, mimeType: exportFormat.mimeType };
  }
  const res = await drive.files.get(
    { fileId, alt: "media" },
    { responseType: "text" }
  );
  return { content: res.data as string, mimeType };
}

export function createServer(clients: DriveClients, rootFolderId?: string): Server {
  const { drive, docs, sheets } = clients;

  const server = new Server(
    { name: "gdrive-sa", version: "1.0.0" },
    { capabilities: { resources: {}, tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: "search",
        description: "Search for files in Google Drive",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Search query (Drive query syntax supported)" },
            pageSize: { type: "number", description: "Max results (default 10, max 100)" },
          },
          required: ["query"],
        },
      },
      {
        name: "list_folder",
        description: "List files in a folder (defaults to root folder if configured)",
        inputSchema: {
          type: "object",
          properties: {
            folderId: { type: "string", description: "Folder ID (omit to use GDRIVE_ROOT_FOLDER_ID)" },
            pageSize: { type: "number", description: "Max results (default 20)" },
          },
        },
      },
      {
        name: "read_file",
        description: "Read the content of a file by ID",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "Google Drive file ID" },
          },
          required: ["fileId"],
        },
      },
      {
        name: "get_file_info",
        description: "Get metadata about a file",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "Google Drive file ID" },
          },
          required: ["fileId"],
        },
      },
      {
        name: "update_file",
        description: "Update the content of a plain-text or binary Drive file",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "Google Drive file ID" },
            content: { type: "string", description: "New file content" },
            mimeType: { type: "string", description: "MIME type (defaults to the file's existing type)" },
          },
          required: ["fileId", "content"],
        },
      },
      {
        name: "update_doc",
        description: "Update a Google Doc. Supply 'content' (markdown supported: # ## ### for headings, **bold**, *italic*) to replace the full document with format-mapped content, or 'find'+'replaceWith' for targeted find-and-replace (preserves existing paragraph style).",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "Google Doc file ID" },
            content: { type: "string", description: "New full document text (plain text or markdown)" },
            find: { type: "string", description: "Text to search for (used with replaceWith)" },
            replaceWith: { type: "string", description: "Replacement text (used with find)" },
            matchCase: { type: "boolean", description: "Case-sensitive find (default true)" },
          },
          required: ["fileId"],
        },
      },
      {
        name: "update_sheet",
        description: "Update a range of cells in a Google Sheet",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "Google Sheets file ID" },
            range: { type: "string", description: "A1 notation range, e.g. 'Sheet1!A1:C3'" },
            values: {
              type: "array",
              description: "2D array of cell values",
              items: { type: "array", items: { type: "string" } },
            },
            valueInputOption: {
              type: "string",
              enum: ["RAW", "USER_ENTERED"],
              description: "How to interpret input values (default: USER_ENTERED)",
            },
          },
          required: ["fileId", "range", "values"],
        },
      },
      {
        name: "list_comments",
        description: "List all comments (and their replies) on a Google Doc",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "Google Drive file ID" },
            includeDeleted: { type: "boolean", description: "Include deleted comments (default false)" },
          },
          required: ["fileId"],
        },
      },
      {
        name: "add_comment",
        description: "Add a comment to a Google Doc, optionally anchored to a specific quoted passage",
        inputSchema: {
          type: "object",
          properties: {
            fileId: { type: "string", description: "Google Drive file ID" },
            content: { type: "string", description: "Comment text" },
            quotedText: { type: "string", description: "Passage in the document to anchor the comment to" },
          },
          required: ["fileId", "content"],
        },
      },
    ],
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    if (name === "search") {
      const { query, pageSize = 10 } = args as { query: string; pageSize?: number };
      const res = await drive.files.list({
        q: query,
        pageSize: Math.min(pageSize, 100),
        fields: "files(id,name,mimeType,modifiedTime,size,webViewLink)",
      });
      const files = res.data.files ?? [];
      return { content: [{ type: "text", text: JSON.stringify(files, null, 2) }] };
    }

    if (name === "list_folder") {
      const { folderId, pageSize = 20 } = args as { folderId?: string; pageSize?: number };
      const id = folderId ?? rootFolderId;
      if (!id) {
        return {
          content: [{ type: "text", text: "No folderId provided and GDRIVE_ROOT_FOLDER_ID is not set" }],
          isError: true,
        };
      }
      const res = await drive.files.list({
        q: `'${id}' in parents and trashed = false`,
        pageSize: Math.min(pageSize, 100),
        fields: "files(id,name,mimeType,modifiedTime,size,webViewLink)",
      });
      const files = res.data.files ?? [];
      return { content: [{ type: "text", text: JSON.stringify(files, null, 2) }] };
    }

    if (name === "read_file") {
      const { fileId } = args as { fileId: string };
      const meta = await drive.files.get({ fileId, fields: "id,name,mimeType" });
      const mimeType = meta.data.mimeType ?? "application/octet-stream";
      const { content, mimeType: resultMime } = await getFileContent(drive, fileId, mimeType);
      const isText = resultMime.startsWith("text/") || resultMime === "application/json";
      return {
        content: [
          isText
            ? { type: "text" as const, text: content }
            : {
                type: "resource" as const,
                resource: {
                  uri: `gdrive:///${fileId}`,
                  mimeType: resultMime,
                  blob: Buffer.from(content).toString("base64"),
                },
              },
        ],
      };
    }

    if (name === "get_file_info") {
      const { fileId } = args as { fileId: string };
      const res = await drive.files.get({
        fileId,
        fields: "id,name,mimeType,modifiedTime,createdTime,size,webViewLink,owners,shared,parents",
      });
      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }

    if (name === "update_file") {
      const { fileId, content, mimeType: contentMime } = args as {
        fileId: string;
        content: string;
        mimeType?: string;
      };
      const meta = await drive.files.get({ fileId, fields: "mimeType" });
      const fileMime = contentMime ?? meta.data.mimeType ?? "text/plain";
      await drive.files.update({
        fileId,
        requestBody: {},
        media: { mimeType: fileMime, body: content },
      });
      return { content: [{ type: "text", text: `File ${fileId} updated successfully` }] };
    }

    if (name === "update_doc") {
      const { fileId, content, find, replaceWith, matchCase = true } = args as {
        fileId: string;
        content?: string;
        find?: string;
        replaceWith?: string;
        matchCase?: boolean;
      };
      if (find !== undefined && replaceWith !== undefined) {
        await docs.documents.batchUpdate({
          documentId: fileId,
          requestBody: {
            requests: [{ replaceAllText: { containsText: { text: find, matchCase }, replaceText: replaceWith } }],
          },
        });
        return { content: [{ type: "text", text: `Replaced "${find}" with "${replaceWith}" in document ${fileId}` }] };
      }
      if (content !== undefined) {
        const doc = await docs.documents.get({ documentId: fileId });
        const bodyContent = doc.data.body?.content ?? [];
        const bodyEndIndex = bodyContent.at(-1)?.endIndex ?? 2;
        const paragraphs = parseMarkdownParagraphs(content);
        const requests = buildMarkdownRequests(paragraphs, bodyEndIndex);
        await docs.documents.batchUpdate({ documentId: fileId, requestBody: { requests } });
        return { content: [{ type: "text", text: `Document ${fileId} updated with formatted content` }] };
      }
      return {
        content: [{ type: "text", text: "Provide 'content' for full replacement or 'find'+'replaceWith' for targeted edit" }],
        isError: true,
      };
    }

    if (name === "update_sheet") {
      const { fileId, range, values, valueInputOption = "USER_ENTERED" } = args as {
        fileId: string;
        range: string;
        values: string[][];
        valueInputOption?: string;
      };
      await sheets.spreadsheets.values.update({
        spreadsheetId: fileId,
        range,
        valueInputOption,
        requestBody: { values },
      });
      return { content: [{ type: "text", text: `Range ${range} in spreadsheet ${fileId} updated` }] };
    }

    if (name === "list_comments") {
      const { fileId, includeDeleted = false } = args as { fileId: string; includeDeleted?: boolean };
      const res = await drive.comments.list({
        fileId,
        includeDeleted,
        fields: "comments(id,content,author,createdTime,modifiedTime,resolved,quotedFileContent,replies(id,content,author,createdTime))",
      });
      return { content: [{ type: "text", text: JSON.stringify(res.data.comments ?? [], null, 2) }] };
    }

    if (name === "add_comment") {
      const { fileId, content, quotedText } = args as { fileId: string; content: string; quotedText?: string };
      const requestBody: drive_v3.Schema$Comment = { content };
      if (quotedText) {
        requestBody.quotedFileContent = { mimeType: "text/plain", value: quotedText };
      }
      const res = await drive.comments.create({
        fileId,
        requestBody,
        fields: "id,content,author,createdTime,quotedFileContent",
      });
      return { content: [{ type: "text", text: JSON.stringify(res.data, null, 2) }] };
    }

    return { content: [{ type: "text", text: `Unknown tool: ${name}` }], isError: true };
  });

  server.setRequestHandler(ListResourcesRequestSchema, async () => {
    if (!rootFolderId) return { resources: [] };
    const res = await drive.files.list({
      q: `'${rootFolderId}' in parents and trashed = false`,
      pageSize: 50,
      fields: "files(id,name,mimeType)",
    });
    const files: drive_v3.Schema$File[] = res.data.files ?? [];
    return {
      resources: files.map((f) => ({
        uri: `gdrive:///${f.id}`,
        mimeType: f.mimeType ?? "application/octet-stream",
        name: f.name ?? f.id ?? "unknown",
      })),
    };
  });

  server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    const fileId = request.params.uri.replace("gdrive:///", "");
    const meta = await drive.files.get({ fileId, fields: "mimeType" });
    const mimeType = meta.data.mimeType ?? "application/octet-stream";
    const { content, mimeType: resultMime } = await getFileContent(drive, fileId, mimeType);
    return {
      contents: [{ uri: request.params.uri, mimeType: resultMime, text: content }],
    };
  });

  return server;
}
