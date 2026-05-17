import { describe, it, expect, vi, beforeEach } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer, getFileContent, EXPORT_MIME_TYPES, type DriveClients } from "../server.js";
import type { drive_v3 } from "googleapis";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function createMockClients() {
  return {
    drive: {
      files: {
        list: vi.fn(),
        get: vi.fn(),
        export: vi.fn(),
        update: vi.fn(),
      },
    },
    docs: {
      documents: {
        batchUpdate: vi.fn(),
      },
    },
    sheets: {
      spreadsheets: {
        values: {
          update: vi.fn(),
        },
      },
    },
  };
}

type MockClients = ReturnType<typeof createMockClients>;

async function connect(mockClients: MockClients, rootFolderId?: string): Promise<Client> {
  const server = createServer(mockClients as unknown as DriveClients, rootFolderId);
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test-client", version: "1.0.0" }, {});
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  return client;
}

type ContentItem = { type: string; text?: string; resource?: { uri: string; mimeType: string; blob?: string } };

function textOf(result: Awaited<ReturnType<Client["callTool"]>>): string {
  const content = result.content as ContentItem[];
  const item = content[0];
  if (item.type !== "text") throw new Error(`Expected text content, got ${item.type}`);
  return item.text!;
}

// ---------------------------------------------------------------------------
// getFileContent — unit tests (no MCP transport needed)
// ---------------------------------------------------------------------------

describe("getFileContent", () => {
  let mockClients: MockClients;

  beforeEach(() => {
    mockClients = createMockClients();
  });

  it("exports Google Doc as markdown", async () => {
    mockClients.drive.files.export.mockResolvedValue({ data: "# Hello" });

    const result = await getFileContent(
      mockClients.drive as unknown as drive_v3.Drive,
      "doc-id",
      "application/vnd.google-apps.document"
    );

    expect(mockClients.drive.files.export).toHaveBeenCalledWith(
      { fileId: "doc-id", mimeType: "text/markdown" },
      { responseType: "text" }
    );
    expect(result).toEqual({ content: "# Hello", mimeType: "text/markdown" });
  });

  it("exports spreadsheet as CSV", async () => {
    mockClients.drive.files.export.mockResolvedValue({ data: "a,b\n1,2" });

    const result = await getFileContent(
      mockClients.drive as unknown as drive_v3.Drive,
      "sheet-id",
      "application/vnd.google-apps.spreadsheet"
    );

    expect(mockClients.drive.files.export).toHaveBeenCalledWith(
      { fileId: "sheet-id", mimeType: "text/csv" },
      { responseType: "text" }
    );
    expect(result).toEqual({ content: "a,b\n1,2", mimeType: "text/csv" });
  });

  it("exports presentation as plain text", async () => {
    mockClients.drive.files.export.mockResolvedValue({ data: "slide text" });

    const result = await getFileContent(
      mockClients.drive as unknown as drive_v3.Drive,
      "pres-id",
      "application/vnd.google-apps.presentation"
    );

    expect(result).toEqual({ content: "slide text", mimeType: "text/plain" });
  });

  it("exports drawing as PNG", async () => {
    mockClients.drive.files.export.mockResolvedValue({ data: "binary-png-data" });

    const result = await getFileContent(
      mockClients.drive as unknown as drive_v3.Drive,
      "draw-id",
      "application/vnd.google-apps.drawing"
    );

    expect(result).toEqual({ content: "binary-png-data", mimeType: "image/png" });
  });

  it("fetches regular text file with alt=media", async () => {
    mockClients.drive.files.get.mockResolvedValue({ data: "plain text content" });

    const result = await getFileContent(
      mockClients.drive as unknown as drive_v3.Drive,
      "file-id",
      "text/plain"
    );

    expect(mockClients.drive.files.get).toHaveBeenCalledWith(
      { fileId: "file-id", alt: "media" },
      { responseType: "text" }
    );
    expect(result).toEqual({ content: "plain text content", mimeType: "text/plain" });
  });

  it("fetches unknown binary file with alt=media, preserving mimeType", async () => {
    mockClients.drive.files.get.mockResolvedValue({ data: "raw-bytes" });

    const result = await getFileContent(
      mockClients.drive as unknown as drive_v3.Drive,
      "bin-id",
      "application/octet-stream"
    );

    expect(result).toEqual({ content: "raw-bytes", mimeType: "application/octet-stream" });
  });
});

// ---------------------------------------------------------------------------
// EXPORT_MIME_TYPES — sanity check
// ---------------------------------------------------------------------------

describe("EXPORT_MIME_TYPES", () => {
  it("contains all four Google Workspace types", () => {
    expect(EXPORT_MIME_TYPES).toHaveProperty("application/vnd.google-apps.document");
    expect(EXPORT_MIME_TYPES).toHaveProperty("application/vnd.google-apps.spreadsheet");
    expect(EXPORT_MIME_TYPES).toHaveProperty("application/vnd.google-apps.presentation");
    expect(EXPORT_MIME_TYPES).toHaveProperty("application/vnd.google-apps.drawing");
  });
});

// ---------------------------------------------------------------------------
// listTools
// ---------------------------------------------------------------------------

describe("listTools", () => {
  it("returns the seven expected tools", async () => {
    const mockClients = createMockClients();
    const client = await connect(mockClients);

    const { tools } = await client.listTools();

    expect(tools).toHaveLength(7);
    const names = tools.map((t) => t.name);
    expect(names).toContain("search");
    expect(names).toContain("list_folder");
    expect(names).toContain("read_file");
    expect(names).toContain("get_file_info");
    expect(names).toContain("update_file");
    expect(names).toContain("update_doc");
    expect(names).toContain("update_sheet");

    await client.close();
  });

  it("search tool has required query parameter", async () => {
    const mockClients = createMockClients();
    const client = await connect(mockClients);

    const { tools } = await client.listTools();
    const search = tools.find((t) => t.name === "search")!;

    expect(search.inputSchema.required).toContain("query");

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// search tool
// ---------------------------------------------------------------------------

describe("search tool", () => {
  it("returns matching files as JSON", async () => {
    const mockClients = createMockClients();
    const files = [
      { id: "1", name: "report.gdoc", mimeType: "application/vnd.google-apps.document" },
      { id: "2", name: "budget.gsheet", mimeType: "application/vnd.google-apps.spreadsheet" },
    ];
    mockClients.drive.files.list.mockResolvedValue({ data: { files } });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "search", arguments: { query: "name contains 'report'" } });

    expect(result.isError).toBeFalsy();
    expect(mockClients.drive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ q: "name contains 'report'", pageSize: 10 })
    );
    expect(JSON.parse(textOf(result))).toEqual(files);

    await client.close();
  });

  it("returns empty array when no files match", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "search", arguments: { query: "nonexistent" } });

    expect(JSON.parse(textOf(result))).toEqual([]);

    await client.close();
  });

  it("caps pageSize at 100", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockClients);
    await client.callTool({ name: "search", arguments: { query: "test", pageSize: 500 } });

    expect(mockClients.drive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ pageSize: 100 })
    );

    await client.close();
  });

  it("handles missing files field in API response", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.list.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "search", arguments: { query: "test" } });

    expect(JSON.parse(textOf(result))).toEqual([]);

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// list_folder tool
// ---------------------------------------------------------------------------

describe("list_folder tool", () => {
  it("lists files in the given folderId", async () => {
    const mockClients = createMockClients();
    const files = [{ id: "f1", name: "file.txt", mimeType: "text/plain" }];
    mockClients.drive.files.list.mockResolvedValue({ data: { files } });

    const client = await connect(mockClients);
    const result = await client.callTool({
      name: "list_folder",
      arguments: { folderId: "folder-abc" },
    });

    expect(mockClients.drive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ q: "'folder-abc' in parents and trashed = false" })
    );
    expect(JSON.parse(textOf(result))).toEqual(files);

    await client.close();
  });

  it("falls back to rootFolderId when folderId is omitted", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockClients, "root-folder-id");
    await client.callTool({ name: "list_folder", arguments: {} });

    expect(mockClients.drive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ q: "'root-folder-id' in parents and trashed = false" })
    );

    await client.close();
  });

  it("returns error when no folderId and no rootFolderId", async () => {
    const mockClients = createMockClients();
    const client = await connect(mockClients);

    const result = await client.callTool({ name: "list_folder", arguments: {} });

    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("GDRIVE_ROOT_FOLDER_ID");

    await client.close();
  });

  it("explicit folderId takes precedence over rootFolderId", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockClients, "root-folder-id");
    await client.callTool({ name: "list_folder", arguments: { folderId: "explicit-folder" } });

    expect(mockClients.drive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ q: "'explicit-folder' in parents and trashed = false" })
    );

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// read_file tool
// ---------------------------------------------------------------------------

describe("read_file tool", () => {
  it("returns text content for a Google Doc", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get.mockResolvedValue({
      data: { id: "doc-id", name: "My Doc", mimeType: "application/vnd.google-apps.document" },
    });
    mockClients.drive.files.export.mockResolvedValue({ data: "# My Document\n\nHello world" });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "doc-id" } });

    expect(result.isError).toBeFalsy();
    expect((result.content as ContentItem[])[0].type).toBe("text");
    expect(textOf(result)).toBe("# My Document\n\nHello world");

    await client.close();
  });

  it("returns text content for a plain text file", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get
      .mockResolvedValueOnce({ data: { mimeType: "text/plain" } })
      .mockResolvedValueOnce({ data: "file content here" });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "txt-id" } });

    expect((result.content as ContentItem[])[0].type).toBe("text");
    expect(textOf(result)).toBe("file content here");

    await client.close();
  });

  it("returns resource blob for a binary/image file", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get.mockResolvedValueOnce({
      data: { mimeType: "application/vnd.google-apps.drawing" },
    });
    mockClients.drive.files.export.mockResolvedValue({ data: "png-binary" });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "draw-id" } });

    const item = (result.content as ContentItem[])[0] as {
      type: "resource";
      resource: { uri: string; mimeType: string; blob: string };
    };
    expect(item.type).toBe("resource");
    expect(item.resource.uri).toBe("gdrive:///draw-id");
    expect(item.resource.mimeType).toBe("image/png");
    expect(item.resource.blob).toBe(Buffer.from("png-binary").toString("base64"));

    await client.close();
  });

  it("uses application/octet-stream when mimeType is missing", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get
      .mockResolvedValueOnce({ data: {} })
      .mockResolvedValueOnce({ data: "raw" });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "no-mime-id" } });

    const item = (result.content as ContentItem[])[0] as { type: string; resource?: { mimeType: string } };
    expect(item.type).toBe("resource");
    expect(item.resource?.mimeType).toBe("application/octet-stream");

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// get_file_info tool
// ---------------------------------------------------------------------------

describe("get_file_info tool", () => {
  it("returns file metadata as JSON", async () => {
    const mockClients = createMockClients();
    const metadata = {
      id: "file-id",
      name: "Report.docx",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      modifiedTime: "2024-01-15T10:00:00Z",
      size: "12345",
    };
    mockClients.drive.files.get.mockResolvedValue({ data: metadata });

    const client = await connect(mockClients);
    const result = await client.callTool({ name: "get_file_info", arguments: { fileId: "file-id" } });

    expect(result.isError).toBeFalsy();
    expect(mockClients.drive.files.get).toHaveBeenCalledWith(
      expect.objectContaining({ fileId: "file-id", fields: expect.stringContaining("id") })
    );
    expect(JSON.parse(textOf(result))).toEqual(metadata);

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// update_file tool
// ---------------------------------------------------------------------------

describe("update_file tool", () => {
  it("updates a plain text file using its existing mimeType", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get.mockResolvedValue({ data: { mimeType: "text/plain" } });
    mockClients.drive.files.update.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    const result = await client.callTool({
      name: "update_file",
      arguments: { fileId: "file-id", content: "new content" },
    });

    expect(result.isError).toBeFalsy();
    expect(mockClients.drive.files.update).toHaveBeenCalledWith(
      expect.objectContaining({
        fileId: "file-id",
        media: expect.objectContaining({ mimeType: "text/plain", body: "new content" }),
      })
    );
    expect(textOf(result)).toContain("file-id");

    await client.close();
  });

  it("respects an explicit mimeType override", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get.mockResolvedValue({ data: { mimeType: "text/plain" } });
    mockClients.drive.files.update.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    await client.callTool({
      name: "update_file",
      arguments: { fileId: "file-id", content: "<p>hi</p>", mimeType: "text/html" },
    });

    expect(mockClients.drive.files.update).toHaveBeenCalledWith(
      expect.objectContaining({
        media: expect.objectContaining({ mimeType: "text/html" }),
      })
    );

    await client.close();
  });

  it("falls back to text/plain when file has no mimeType", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get.mockResolvedValue({ data: {} });
    mockClients.drive.files.update.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    await client.callTool({
      name: "update_file",
      arguments: { fileId: "file-id", content: "data" },
    });

    expect(mockClients.drive.files.update).toHaveBeenCalledWith(
      expect.objectContaining({
        media: expect.objectContaining({ mimeType: "text/plain" }),
      })
    );

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// update_doc tool
// ---------------------------------------------------------------------------

describe("update_doc tool", () => {
  it("replaces full document content via Drive API when 'content' is provided", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.update.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    const result = await client.callTool({
      name: "update_doc",
      arguments: { fileId: "doc-id", content: "# New content" },
    });

    expect(result.isError).toBeFalsy();
    expect(mockClients.drive.files.update).toHaveBeenCalledWith(
      expect.objectContaining({
        fileId: "doc-id",
        media: expect.objectContaining({ mimeType: "text/plain", body: "# New content" }),
      })
    );

    await client.close();
  });

  it("uses Docs API batchUpdate for find-and-replace", async () => {
    const mockClients = createMockClients();
    mockClients.docs.documents.batchUpdate.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    const result = await client.callTool({
      name: "update_doc",
      arguments: { fileId: "doc-id", find: "Hello", replaceWith: "Hi" },
    });

    expect(result.isError).toBeFalsy();
    expect(mockClients.docs.documents.batchUpdate).toHaveBeenCalledWith(
      expect.objectContaining({
        documentId: "doc-id",
        requestBody: {
          requests: [
            expect.objectContaining({
              replaceAllText: expect.objectContaining({
                containsText: { text: "Hello", matchCase: true },
                replaceText: "Hi",
              }),
            }),
          ],
        },
      })
    );
    expect(textOf(result)).toContain("Hello");

    await client.close();
  });

  it("respects matchCase: false", async () => {
    const mockClients = createMockClients();
    mockClients.docs.documents.batchUpdate.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    await client.callTool({
      name: "update_doc",
      arguments: { fileId: "doc-id", find: "hello", replaceWith: "hi", matchCase: false },
    });

    expect(mockClients.docs.documents.batchUpdate).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: {
          requests: [
            expect.objectContaining({
              replaceAllText: expect.objectContaining({
                containsText: { text: "hello", matchCase: false },
              }),
            }),
          ],
        },
      })
    );

    await client.close();
  });

  it("returns error when neither content nor find/replaceWith are provided", async () => {
    const mockClients = createMockClients();
    const client = await connect(mockClients);

    const result = await client.callTool({
      name: "update_doc",
      arguments: { fileId: "doc-id" },
    });

    expect(result.isError).toBe(true);

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// update_sheet tool
// ---------------------------------------------------------------------------

describe("update_sheet tool", () => {
  it("updates a cell range with USER_ENTERED by default", async () => {
    const mockClients = createMockClients();
    mockClients.sheets.spreadsheets.values.update.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    const values = [["a", "b"], ["1", "2"]];
    const result = await client.callTool({
      name: "update_sheet",
      arguments: { fileId: "sheet-id", range: "Sheet1!A1:B2", values },
    });

    expect(result.isError).toBeFalsy();
    expect(mockClients.sheets.spreadsheets.values.update).toHaveBeenCalledWith(
      expect.objectContaining({
        spreadsheetId: "sheet-id",
        range: "Sheet1!A1:B2",
        valueInputOption: "USER_ENTERED",
        requestBody: { values },
      })
    );

    await client.close();
  });

  it("passes RAW valueInputOption when specified", async () => {
    const mockClients = createMockClients();
    mockClients.sheets.spreadsheets.values.update.mockResolvedValue({ data: {} });

    const client = await connect(mockClients);
    await client.callTool({
      name: "update_sheet",
      arguments: { fileId: "sheet-id", range: "A1", values: [["=SUM(1,2)"]], valueInputOption: "RAW" },
    });

    expect(mockClients.sheets.spreadsheets.values.update).toHaveBeenCalledWith(
      expect.objectContaining({ valueInputOption: "RAW" })
    );

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// unknown tool
// ---------------------------------------------------------------------------

describe("unknown tool", () => {
  it("returns isError with message for unrecognized tool name", async () => {
    const mockClients = createMockClients();
    const client = await connect(mockClients);

    const result = await client.callTool({ name: "does_not_exist", arguments: {} });

    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("does_not_exist");

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// listResources
// ---------------------------------------------------------------------------

describe("listResources", () => {
  it("returns resources from the root folder", async () => {
    const mockClients = createMockClients();
    const files = [
      { id: "f1", name: "doc.gdoc", mimeType: "application/vnd.google-apps.document" },
      { id: "f2", name: "sheet.gsheet", mimeType: "application/vnd.google-apps.spreadsheet" },
    ];
    mockClients.drive.files.list.mockResolvedValue({ data: { files } });

    const client = await connect(mockClients, "root-id");
    const { resources } = await client.listResources();

    expect(resources).toHaveLength(2);
    expect(resources[0]).toMatchObject({ uri: "gdrive:///f1", name: "doc.gdoc" });
    expect(resources[1]).toMatchObject({ uri: "gdrive:///f2", name: "sheet.gsheet" });

    await client.close();
  });

  it("returns empty list when rootFolderId is not configured", async () => {
    const mockClients = createMockClients();
    const client = await connect(mockClients);

    const { resources } = await client.listResources();

    expect(resources).toHaveLength(0);
    expect(mockClients.drive.files.list).not.toHaveBeenCalled();

    await client.close();
  });

  it("falls back to file id as name when name is missing", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.list.mockResolvedValue({
      data: { files: [{ id: "f3", mimeType: "text/plain" }] },
    });

    const client = await connect(mockClients, "root-id");
    const { resources } = await client.listResources();

    expect(resources[0].name).toBe("f3");

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// readResource
// ---------------------------------------------------------------------------

describe("readResource", () => {
  it("extracts fileId from URI and returns content", async () => {
    const mockClients = createMockClients();
    mockClients.drive.files.get
      .mockResolvedValueOnce({ data: { mimeType: "text/plain" } })
      .mockResolvedValueOnce({ data: "resource text content" });

    const client = await connect(mockClients, "root-id");
    mockClients.drive.files.list.mockResolvedValue({
      data: { files: [{ id: "res-file-id", name: "note.txt", mimeType: "text/plain" }] },
    });
    await client.listResources();

    const { contents } = await client.readResource({ uri: "gdrive:///res-file-id" });

    expect(contents).toHaveLength(1);
    expect(contents[0].uri).toBe("gdrive:///res-file-id");
    expect(contents[0].mimeType).toBe("text/plain");
    expect((contents[0] as { uri: string; mimeType: string; text: string }).text).toBe(
      "resource text content"
    );

    await client.close();
  });
});
