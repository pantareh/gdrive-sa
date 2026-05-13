import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { drive_v3 } from "googleapis";

export const EXPORT_MIME_TYPES: Record<string, { mimeType: string; ext: string }> = {
  "application/vnd.google-apps.document": { mimeType: "text/markdown", ext: "md" },
  "application/vnd.google-apps.spreadsheet": { mimeType: "text/csv", ext: "csv" },
  "application/vnd.google-apps.presentation": { mimeType: "text/plain", ext: "txt" },
  "application/vnd.google-apps.drawing": { mimeType: "image/png", ext: "png" },
};

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

export function createServer(drive: drive_v3.Drive, rootFolderId?: string): Server {
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
