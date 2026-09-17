import { createMcpHandler } from "mcp-handler";
import { withMcpAuth } from "mcp-handler";
import { z } from "zod";
import { auth } from "@/app/auth";
import { google } from "googleapis";

async function getDriveClient(accessToken: string) {
  const oauth2Client = new google.auth.OAuth2();
  oauth2Client.setCredentials({ access_token: accessToken });
  return google.drive({ version: "v3", auth: oauth2Client });
}

async function getDocsClient(accessToken: string) {
  const oauth2Client = new google.auth.OAuth2();
  oauth2Client.setCredentials({ access_token: accessToken });
  return google.docs({ version: "v1", auth: oauth2Client });
}

async function getSheetsClient(accessToken: string) {
  const oauth2Client = new google.auth.OAuth2();
  oauth2Client.setCredentials({ access_token: accessToken });
  return google.sheets({ version: "v4", auth: oauth2Client });
}

const handler = createMcpHandler((server) => {
  server.tool(
    "list_drive_files",
    "List files in the user's Google Drive, optionally filtered by a search query (Drive query syntax) or folder ID.",
    {
      query: z.string().optional().describe("Optional Drive search query, e.g. \"name contains 'PE'\""),
      folderId: z.string().optional().describe("Optional parent folder ID to list contents of"),
      pageSize: z.number().min(1).max(100).default(20),
    },
    async ({ query, folderId, pageSize }, { authInfo }) => {
      const accessToken = authInfo?.extra?.googleAccessToken as string;
      const drive = await getDriveClient(accessToken);
      const qParts: string[] = ["trashed = false"];
      if (query) qParts.push(query);
      if (folderId) qParts.push(`'${folderId}' in parents`);
      const res = await drive.files.list({
        q: qParts.join(" and "),
        pageSize,
        fields: "files(id, name, mimeType, modifiedTime, webViewLink, parents)",
      });
      return {
        content: [{ type: "text", text: JSON.stringify(res.data.files, null, 2) }],
      };
    }
  );

  server.tool(
    "read_drive_file",
    "Read the content of a Google Drive file. Works for Google Docs, Sheets, and plain text files.",
    {
      fileId: z.string().describe("The Google Drive file ID"),
    },
    async ({ fileId }, { authInfo }) => {
      const accessToken = authInfo?.extra?.googleAccessToken as string;
      const drive = await getDriveClient(accessToken);
      const meta = await drive.files.get({ fileId, fields: "mimeType, name" });
      const mimeType = meta.data.mimeType;

      if (mimeType === "application/vnd.google-apps.document") {
        const docs = await getDocsClient(accessToken);
        const doc = await docs.documents.get({ documentId: fileId });
        const text = (doc.data.body?.content || [])
          .flatMap((el) => el.paragraph?.elements || [])
          .map((el) => el.textRun?.content || "")
          .join("");
        return { content: [{ type: "text", text }] };
      }

      if (mimeType === "application/vnd.google-apps.spreadsheet") {
        const sheets = await getSheetsClient(accessToken);
        const values = await sheets.spreadsheets.values.get({
          spreadsheetId: fileId,
          range: "A1:Z1000",
        });
        return { content: [{ type: "text", text: JSON.stringify(values.data.values, null, 2) }] };
      }

      const res = await drive.files.get({ fileId, alt: "media" }, { responseType: "text" });
      return { content: [{ type: "text", text: res.data as string }] };
    }
  );

  server.tool(
    "create_drive_doc",
    "Create a new Google Doc with the given title and initial text content.",
    {
      title: z.string(),
      content: z.string().optional(),
      folderId: z.string().optional().describe("Optional parent folder ID"),
    },
    async ({ title, content, folderId }, { authInfo }) => {
      const accessToken = authInfo?.extra?.googleAccessToken as string;
      const docs = await getDocsClient(accessToken);
      const drive = await getDriveClient(accessToken);
      const doc = await docs.documents.create({ requestBody: { title } });
      const documentId = doc.data.documentId!;

      if (content) {
        await docs.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [{ insertText: { location: { index: 1 }, text: content } }],
          },
        });
      }

      if (folderId) {
        await drive.files.update({
          fileId: documentId,
          addParents: folderId,
          fields: "id, parents",
        });
      }

      return {
        content: [
          {
            type: "text",
            text: `Created Google Doc "${title}" (ID: ${documentId}) — https://docs.google.com/document/d/${documentId}/edit`,
          },
        ],
      };
    }
  );

  server.tool(
    "create_drive_sheet",
    "Create a new Google Sheet with the given title and optional initial rows (array of arrays).",
    {
      title: z.string(),
      rows: z.array(z.array(z.string())).optional().describe("Initial rows, e.g. [[\"Header1\",\"Header2\"],[\"a\",\"b\"]]"),
      folderId: z.string().optional(),
    },
    async ({ title, rows, folderId }, { authInfo }) => {
      const accessToken = authInfo?.extra?.googleAccessToken as string;
      const sheets = await getSheetsClient(accessToken);
      const drive = await getDriveClient(accessToken);
      const sheet = await sheets.spreadsheets.create({ requestBody: { properties: { title } } });
      const spreadsheetId = sheet.data.spreadsheetId!;

      if (rows && rows.length > 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: "A1",
          valueInputOption: "RAW",
          requestBody: { values: rows },
        });
      }

      if (folderId) {
        await drive.files.update({
          fileId: spreadsheetId,
          addParents: folderId,
          fields: "id, parents",
        });
      }

      return {
        content: [
          {
            type: "text",
            text: `Created Google Sheet "${title}" (ID: ${spreadsheetId}) — https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
          },
        ],
      };
    }
  );

  server.tool(
    "update_drive_sheet",
    "Write or overwrite values in a range of an existing Google Sheet.",
    {
      spreadsheetId: z.string(),
      range: z.string().describe("A1 notation range, e.g. \"Sheet1!A1:C10\""),
      values: z.array(z.array(z.string())),
    },
    async ({ spreadsheetId, range, values }, { authInfo }) => {
      const accessToken = authInfo?.extra?.googleAccessToken as string;
      const sheets = await getSheetsClient(accessToken);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: "RAW",
        requestBody: { values },
      });
      return {
        content: [{ type: "text", text: `Updated range ${range} in spreadsheet ${spreadsheetId}.` }],
      };
    }
  );

  server.tool(
    "append_drive_sheet",
    "Append rows to the end of an existing Google Sheet, without overwriting existing data.",
    {
      spreadsheetId: z.string(),
      range: z.string().describe("A1 notation range/sheet name to append after, e.g. \"Sheet1\""),
      values: z.array(z.array(z.string())),
    },
    async ({ spreadsheetId, range, values }, { authInfo }) => {
      const accessToken = authInfo?.extra?.googleAccessToken as string;
      const sheets = await getSheetsClient(accessToken);
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption: "RAW",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values },
      });
      return {
        content: [{ type: "text", text: `Appended ${values.length} row(s) to ${range} in spreadsheet ${spreadsheetId}.` }],
      };
    }
  );

  server.tool(
    "update_drive_doc",
    "Append text content to the end of an existing Google Doc.",
    {
      documentId: z.string(),
      text: z.string(),
    },
    async ({ documentId, text }, { authInfo }) => {
      const accessToken = authInfo?.extra?.googleAccessToken as string;
      const docs = await getDocsClient(accessToken);
      const doc = await docs.documents.get({ documentId });
      const endIndex = doc.data.body?.content?.slice(-1)[0]?.endIndex || 1;
      await docs.documents.batchUpdate({
        documentId,
        requestBody: {
          requests: [{ insertText: { location: { index: endIndex - 1 }, text } }],
        },
      });
      return {
        content: [{ type: "text", text: `Appended text to document ${documentId}.` }],
      };
    }
  );
});

// Wrap the handler with MCP auth, extracting the Google access token
// from the NextAuth session so tools can call Google APIs on the user's behalf.
const authHandler = withMcpAuth(
  handler,
  async (req, bearerToken) => {
    const session = await auth();
    if (!session?.user) return undefined;
    return {
      token: bearerToken || "session",
      clientId: session.user.email || "unknown",
      scopes: [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents",
        "https://www.googleapis.com/auth/spreadsheets",
      ],
      extra: {
        googleAccessToken: (session as any).googleAccessToken,
      },
    };
  },
  { required: true }
);

export { authHandler as GET, authHandler as POST, authHandler as DELETE };
