import { describe, it, expect, vi, beforeEach } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer, getFileContent, EXPORT_MIME_TYPES } from "../server.js";
import type { drive_v3 } from "googleapis";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function createMockDrive() {
  return {
    files: {
      list: vi.fn(),
      get: vi.fn(),
      export: vi.fn(),
    },
  };
}

type MockDrive = ReturnType<typeof createMockDrive>;

async function connect(mockDrive: MockDrive, rootFolderId?: string): Promise<Client> {
  const server = createServer(mockDrive as unknown as drive_v3.Drive, rootFolderId);
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test-client", version: "1.0.0" }, {});
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  return client;
}

function textOf(result: Awaited<ReturnType<Client["callTool"]>>): string {
  const item = result.content[0];
  if (item.type !== "text") throw new Error(`Expected text content, got ${item.type}`);
  return (item as { type: "text"; text: string }).text;
}

// ---------------------------------------------------------------------------
// getFileContent — unit tests (no MCP transport needed)
// ---------------------------------------------------------------------------

describe("getFileContent", () => {
  let mockDrive: MockDrive;

  beforeEach(() => {
    mockDrive = createMockDrive();
  });

  it("exports Google Doc as markdown", async () => {
    mockDrive.files.export.mockResolvedValue({ data: "# Hello" });

    const result = await getFileContent(
      mockDrive as unknown as drive_v3.Drive,
      "doc-id",
      "application/vnd.google-apps.document"
    );

    expect(mockDrive.files.export).toHaveBeenCalledWith(
      { fileId: "doc-id", mimeType: "text/markdown" },
      { responseType: "text" }
    );
    expect(result).toEqual({ content: "# Hello", mimeType: "text/markdown" });
  });

  it("exports spreadsheet as CSV", async () => {
    mockDrive.files.export.mockResolvedValue({ data: "a,b\n1,2" });

    const result = await getFileContent(
      mockDrive as unknown as drive_v3.Drive,
      "sheet-id",
      "application/vnd.google-apps.spreadsheet"
    );

    expect(mockDrive.files.export).toHaveBeenCalledWith(
      { fileId: "sheet-id", mimeType: "text/csv" },
      { responseType: "text" }
    );
    expect(result).toEqual({ content: "a,b\n1,2", mimeType: "text/csv" });
  });

  it("exports presentation as plain text", async () => {
    mockDrive.files.export.mockResolvedValue({ data: "slide text" });

    const result = await getFileContent(
      mockDrive as unknown as drive_v3.Drive,
      "pres-id",
      "application/vnd.google-apps.presentation"
    );

    expect(result).toEqual({ content: "slide text", mimeType: "text/plain" });
  });

  it("exports drawing as PNG", async () => {
    mockDrive.files.export.mockResolvedValue({ data: "binary-png-data" });

    const result = await getFileContent(
      mockDrive as unknown as drive_v3.Drive,
      "draw-id",
      "application/vnd.google-apps.drawing"
    );

    expect(result).toEqual({ content: "binary-png-data", mimeType: "image/png" });
  });

  it("fetches regular text file with alt=media", async () => {
    mockDrive.files.get.mockResolvedValue({ data: "plain text content" });

    const result = await getFileContent(
      mockDrive as unknown as drive_v3.Drive,
      "file-id",
      "text/plain"
    );

    expect(mockDrive.files.get).toHaveBeenCalledWith(
      { fileId: "file-id", alt: "media" },
      { responseType: "text" }
    );
    expect(result).toEqual({ content: "plain text content", mimeType: "text/plain" });
  });

  it("fetches unknown binary file with alt=media, preserving mimeType", async () => {
    mockDrive.files.get.mockResolvedValue({ data: "raw-bytes" });

    const result = await getFileContent(
      mockDrive as unknown as drive_v3.Drive,
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
  it("returns the four expected tools", async () => {
    const mockDrive = createMockDrive();
    const client = await connect(mockDrive);

    const { tools } = await client.listTools();

    expect(tools).toHaveLength(4);
    const names = tools.map((t) => t.name);
    expect(names).toContain("search");
    expect(names).toContain("list_folder");
    expect(names).toContain("read_file");
    expect(names).toContain("get_file_info");

    await client.close();
  });

  it("search tool has required query parameter", async () => {
    const mockDrive = createMockDrive();
    const client = await connect(mockDrive);

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
    const mockDrive = createMockDrive();
    const files = [
      { id: "1", name: "report.gdoc", mimeType: "application/vnd.google-apps.document" },
      { id: "2", name: "budget.gsheet", mimeType: "application/vnd.google-apps.spreadsheet" },
    ];
    mockDrive.files.list.mockResolvedValue({ data: { files } });

    const client = await connect(mockDrive);
    const result = await client.callTool({ name: "search", arguments: { query: "name contains 'report'" } });

    expect(result.isError).toBeFalsy();
    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ q: "name contains 'report'", pageSize: 10 })
    );
    expect(JSON.parse(textOf(result))).toEqual(files);

    await client.close();
  });

  it("returns empty array when no files match", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockDrive);
    const result = await client.callTool({ name: "search", arguments: { query: "nonexistent" } });

    expect(JSON.parse(textOf(result))).toEqual([]);

    await client.close();
  });

  it("caps pageSize at 100", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockDrive);
    await client.callTool({ name: "search", arguments: { query: "test", pageSize: 500 } });

    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ pageSize: 100 })
    );

    await client.close();
  });

  it("handles missing files field in API response", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: {} });

    const client = await connect(mockDrive);
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
    const mockDrive = createMockDrive();
    const files = [{ id: "f1", name: "file.txt", mimeType: "text/plain" }];
    mockDrive.files.list.mockResolvedValue({ data: { files } });

    const client = await connect(mockDrive);
    const result = await client.callTool({
      name: "list_folder",
      arguments: { folderId: "folder-abc" },
    });

    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ q: "'folder-abc' in parents and trashed = false" })
    );
    expect(JSON.parse(textOf(result))).toEqual(files);

    await client.close();
  });

  it("falls back to rootFolderId when folderId is omitted", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockDrive, "root-folder-id");
    await client.callTool({ name: "list_folder", arguments: {} });

    expect(mockDrive.files.list).toHaveBeenCalledWith(
      expect.objectContaining({ q: "'root-folder-id' in parents and trashed = false" })
    );

    await client.close();
  });

  it("returns error when no folderId and no rootFolderId", async () => {
    const mockDrive = createMockDrive();
    const client = await connect(mockDrive /* no rootFolderId */);

    const result = await client.callTool({ name: "list_folder", arguments: {} });

    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("GDRIVE_ROOT_FOLDER_ID");

    await client.close();
  });

  it("explicit folderId takes precedence over rootFolderId", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.list.mockResolvedValue({ data: { files: [] } });

    const client = await connect(mockDrive, "root-folder-id");
    await client.callTool({ name: "list_folder", arguments: { folderId: "explicit-folder" } });

    expect(mockDrive.files.list).toHaveBeenCalledWith(
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
    const mockDrive = createMockDrive();
    mockDrive.files.get.mockResolvedValue({
      data: { id: "doc-id", name: "My Doc", mimeType: "application/vnd.google-apps.document" },
    });
    mockDrive.files.export.mockResolvedValue({ data: "# My Document\n\nHello world" });

    const client = await connect(mockDrive);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "doc-id" } });

    expect(result.isError).toBeFalsy();
    expect(result.content[0].type).toBe("text");
    expect(textOf(result)).toBe("# My Document\n\nHello world");

    await client.close();
  });

  it("returns text content for a plain text file", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.get
      .mockResolvedValueOnce({ data: { mimeType: "text/plain" } })  // metadata call
      .mockResolvedValueOnce({ data: "file content here" });          // media call

    const client = await connect(mockDrive);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "txt-id" } });

    expect(result.content[0].type).toBe("text");
    expect(textOf(result)).toBe("file content here");

    await client.close();
  });

  it("returns resource blob for a binary/image file", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.get
      .mockResolvedValueOnce({ data: { mimeType: "application/vnd.google-apps.drawing" } })
      .mockResolvedValue({ data: "png-binary" }); // not called for drawings
    mockDrive.files.export.mockResolvedValue({ data: "png-binary" });

    const client = await connect(mockDrive);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "draw-id" } });

    const item = result.content[0] as {
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
    const mockDrive = createMockDrive();
    mockDrive.files.get
      .mockResolvedValueOnce({ data: {} })                    // no mimeType
      .mockResolvedValueOnce({ data: "raw" });               // media call

    const client = await connect(mockDrive);
    const result = await client.callTool({ name: "read_file", arguments: { fileId: "no-mime-id" } });

    const item = result.content[0] as { type: string; resource?: { mimeType: string } };
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
    const mockDrive = createMockDrive();
    const metadata = {
      id: "file-id",
      name: "Report.docx",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      modifiedTime: "2024-01-15T10:00:00Z",
      size: "12345",
    };
    mockDrive.files.get.mockResolvedValue({ data: metadata });

    const client = await connect(mockDrive);
    const result = await client.callTool({ name: "get_file_info", arguments: { fileId: "file-id" } });

    expect(result.isError).toBeFalsy();
    expect(mockDrive.files.get).toHaveBeenCalledWith(
      expect.objectContaining({ fileId: "file-id", fields: expect.stringContaining("id") })
    );
    expect(JSON.parse(textOf(result))).toEqual(metadata);

    await client.close();
  });
});

// ---------------------------------------------------------------------------
// unknown tool
// ---------------------------------------------------------------------------

describe("unknown tool", () => {
  it("returns isError with message for unrecognized tool name", async () => {
    const mockDrive = createMockDrive();
    const client = await connect(mockDrive);

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
    const mockDrive = createMockDrive();
    const files = [
      { id: "f1", name: "doc.gdoc", mimeType: "application/vnd.google-apps.document" },
      { id: "f2", name: "sheet.gsheet", mimeType: "application/vnd.google-apps.spreadsheet" },
    ];
    mockDrive.files.list.mockResolvedValue({ data: { files } });

    const client = await connect(mockDrive, "root-id");
    const { resources } = await client.listResources();

    expect(resources).toHaveLength(2);
    expect(resources[0]).toMatchObject({ uri: "gdrive:///f1", name: "doc.gdoc" });
    expect(resources[1]).toMatchObject({ uri: "gdrive:///f2", name: "sheet.gsheet" });

    await client.close();
  });

  it("returns empty list when rootFolderId is not configured", async () => {
    const mockDrive = createMockDrive();
    const client = await connect(mockDrive /* no rootFolderId */);

    const { resources } = await client.listResources();

    expect(resources).toHaveLength(0);
    expect(mockDrive.files.list).not.toHaveBeenCalled();

    await client.close();
  });

  it("falls back to file id as name when name is missing", async () => {
    const mockDrive = createMockDrive();
    mockDrive.files.list.mockResolvedValue({
      data: { files: [{ id: "f3", mimeType: "text/plain" /* no name */ }] },
    });

    const client = await connect(mockDrive, "root-id");
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
    const mockDrive = createMockDrive();
    mockDrive.files.get
      .mockResolvedValueOnce({ data: { mimeType: "text/plain" } })
      .mockResolvedValueOnce({ data: "resource text content" });

    const client = await connect(mockDrive, "root-id");
    // listResources so the client knows about the resource
    mockDrive.files.list.mockResolvedValue({
      data: { files: [{ id: "res-file-id", name: "note.txt", mimeType: "text/plain" }] },
    });
    await client.listResources(); // registers the URI

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
