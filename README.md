# @pantareh/mcp-gdrive-sa

An [MCP](https://modelcontextprotocol.io) server for Google Drive that authenticates via a **Service Account** (JSON key file) instead of OAuth. This is useful for server-side deployments, CI environments, and multi-user setups where a shared service account owns or has been granted access to the relevant Drive files.

## Why service account instead of OAuth?

The official `@modelcontextprotocol/server-gdrive` package uses OAuth, which requires an interactive browser-based login and ties credentials to a single user session. A service account credential file is a static JSON key that can be deployed as an environment variable or a secret, with no browser interaction required.

## Prerequisites

1. **Google Cloud project** with the Drive API enabled.
2. A **Service Account** with a downloaded JSON key file.
3. The service account must have access to the files/folders it needs to read. Two ways to grant access:
   - Share individual files or folders with the service account's email address (e.g. `my-sa@my-project.iam.gserviceaccount.com`).
   - Use **Domain-wide delegation** if you need to impersonate a Workspace user (requires Workspace admin setup).

## Installation

```bash
npm install -g @pantareh/mcp-gdrive-sa
```

Or run directly without installing:

```bash
npx @pantareh/mcp-gdrive-sa
```

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `GOOGLE_APPLICATION_CREDENTIALS` | Yes | Absolute path to the service account JSON key file |
| `GDRIVE_ROOT_FOLDER_ID` | No | Default folder ID used by `list_folder` when no `folderId` is passed, and exposed via `listResources` |

## Usage with Claude Desktop

Add this to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "gdrive": {
      "command": "npx",
      "args": ["-y", "@pantareh/mcp-gdrive-sa"],
      "env": {
        "GOOGLE_APPLICATION_CREDENTIALS": "/absolute/path/to/service-account.json",
        "GDRIVE_ROOT_FOLDER_ID": "your-root-folder-id"
      }
    }
  }
}
```

Config file locations:
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

## Available tools

### `search`
Search for files in Drive using [Drive query syntax](https://developers.google.com/drive/api/guides/search-files).

```
query      (required) Drive query string, e.g. "name contains 'report'"
pageSize   (optional) Max results, default 10, capped at 100
```

### `list_folder`
List files inside a folder.

```
folderId   (optional) Drive folder ID; falls back to GDRIVE_ROOT_FOLDER_ID
pageSize   (optional) Max results, default 20
```

### `read_file`
Read a file's content by ID. Google Workspace files are automatically exported:

| Google type | Exported as |
|---|---|
| Document | Markdown (`text/markdown`) |
| Spreadsheet | CSV (`text/csv`) |
| Presentation | Plain text |
| Drawing | PNG (returned as base64 blob) |

All other files are returned as-is.

### `get_file_info`
Get metadata for a file (name, mimeType, size, owners, parents, etc.).

```
fileId     (required) Drive file ID
```

### `update_file`
Update the content of a plain-text or binary Drive file.

```
fileId     (required) Drive file ID
content    (required) New file content
mimeType   (optional) MIME type — defaults to the file's existing type
```

### `update_doc`
Update a Google Doc. Two modes:

- **Full replace** (`content`): replaces the entire document using the Docs API with format-aware conversion. Markdown syntax is mapped to native Google Docs styles:

  | Markdown | Google Docs style |
  |---|---|
  | `# Heading` | Heading 1 |
  | `## Heading` | Heading 2 |
  | `### Heading` | Heading 3 |
  | `**text**` | Bold |
  | `*text*` or `_text_` | Italic |
  | `***text***` | Bold + Italic |
  | Plain text | Normal Text |

- **Find and replace** (`find` + `replaceWith`): replaces all occurrences of `find` in the document. The existing paragraph style (heading, body, etc.) of the matched text is preserved.

```
fileId       (required) Google Doc file ID
content      (optional) New full document text (markdown supported)
find         (optional) Text to search for
replaceWith  (optional) Replacement text
matchCase    (optional) Case-sensitive match, default true
```

### `update_sheet`
Update a range of cells in a Google Sheet.

```
fileId            (required) Google Sheets file ID
range             (required) A1 notation, e.g. "Sheet1!A1:C3"
values            (required) 2D array of cell values
valueInputOption  (optional) "RAW" or "USER_ENTERED" (default)
```

### `list_comments`
List all comments and their replies on a Google Doc.

```
fileId          (required) Google Drive file ID
includeDeleted  (optional) Include deleted comments, default false
```

Each comment includes author, timestamp, content, resolved status, quoted passage, and any replies.

### `add_comment`
Add a comment to a Google Doc. Optionally anchor it to a specific passage.

```
fileId       (required) Google Drive file ID
content      (required) Comment text
quotedText   (optional) Passage in the document to anchor the comment to
```

## Resources

When `GDRIVE_ROOT_FOLDER_ID` is set, the server exposes files in that folder as MCP resources, accessible via URIs of the form `gdrive:///<fileId>`.

## Publishing

The package is published to npm as `@pantareh/mcp-gdrive-sa` with public access (set in `publishConfig`).

```bash
# 1. Make sure you're logged in to npm
npm login

# 2. Build and verify the output
npm run build

# 3. Preview what will be included in the package
npm pack --dry-run

# 4. Bump the version (choose: patch | minor | major)
npm version patch   # e.g. 1.0.0 → 1.0.1

# 5. Publish
npm publish
```

The `prepare` script runs `npm run build` automatically before publishing, so `dist/` is always up to date. Only the `dist/` folder and `README.md` are included in the published package (controlled by the `files` field in `package.json`).

> **Scoped package note:** the first time you publish a scoped package (`@pantareh/...`), npm defaults to private. The `publishConfig.access = "public"` in `package.json` overrides this automatically, so no extra flags are needed.

### Publishing via GitHub Actions (recommended)

A workflow in `.github/workflows/publish.yml` publishes automatically when a version tag is pushed. It uses an `NPM_TOKEN` secret stored in the repo — no local credentials needed.

```bash
npm version patch        # bumps version and creates a git tag
git push --follow-tags   # pushes commits + tag, triggering the workflow
```

> **Note:** pushing commits alone does **not** trigger the workflow — the tag push is what fires it.

## Development

```bash
# Install dependencies
npm install

# Build TypeScript
npm run build

# Run tests (27 unit + integration tests, no network calls)
npm test

# Watch mode
npm run test:watch

# Coverage report
npm run test:coverage
```

### Testing approach

Tests use the MCP SDK's `InMemoryTransport` to wire a real `Server` instance to a real `Client` instance in-process, with the Google Drive API mocked via `vi.fn()`. This exercises the full MCP protocol layer without network calls.
