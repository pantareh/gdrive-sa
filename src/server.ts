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
        description: "Update a Google Doc. Supply 'content' to replace the full document text, or 'find'+'replaceWith' for targeted find-and-replace.",
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
        await drive.files.update({
          fileId,
          media: { mimeType: "text/plain", body: content },
        });
        return { content: [{ type: "text", text: `Document ${fileId} content replaced` }] };
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
