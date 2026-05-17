#!/usr/bin/env node
import { google } from "googleapis";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createServer } from "./server.js";

const CREDENTIAL_PATH = process.env.GOOGLE_APPLICATION_CREDENTIALS;
const ROOT_FOLDER_ID = process.env.GDRIVE_ROOT_FOLDER_ID;

if (!CREDENTIAL_PATH) {
  console.error("GOOGLE_APPLICATION_CREDENTIALS env var is required");
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  keyFile: CREDENTIAL_PATH,
  scopes: ["https://www.googleapis.com/auth/drive"],
});

const drive = google.drive({ version: "v3", auth });
const docs = google.docs({ version: "v1", auth });
const sheets = google.sheets({ version: "v4", auth });
const server = createServer({ drive, docs, sheets }, ROOT_FOLDER_ID);

const transport = new StdioServerTransport();
await server.connect(transport);
