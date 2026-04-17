import { describe, it, afterEach } from "node:test";
import assert from "node:assert/strict";
import express, { type RequestHandler } from "express";
import request from "supertest";
import { createCausationRouter } from "./causation.js";

const authStub: RequestHandler = (req, _res, next) => {
  req.user = {
    userId: "test-user",
    email: "test@example.com",
    isAdmin: false,
  };
  req.integrationAuthMode = "internal_api_key";
  next();
};

function buildApp(customDeps: any) {
  const app = express();
  app.use(express.json({ limit: "2mb" }));
  app.use("/causation", createCausationRouter(authStub, customDeps));
  return app;
}

afterEach(() => {
  global.fetch = originalFetch;
});

describe("causation route with openaiFileIdRefs", () => {
  it("test-openai-file accepts mime_type null when name ends with .pdf", async () => {
    const app = buildApp({});
    const response = await request(app).post("/causation/test-openai-file").send({
      openaiFileIdRefs: [
        {
          name: "DS-1-1570.pdf",
          id: "file-123",
          mime_type: null,
          download_link: "https://files.example.com/download/file-123",
        },
      ],
    });

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(response.body.data.reached_controller, true);
  });

  it("accepts openaiFileIdRefs valid request", async () => {
    global.fetch = async () =>
      new Response(Buffer.from("%PDF-1.4\nchatgpt-file"), {
        status: 200,
        headers: { "content-type": "application/pdf" },
      });

    const deps = {
      readRegistroRows: async () => ({
        spreadsheetId: "sheet-id",
        gid: "42421166",
        rows: [
          {
            rowNumber: 7,
            dateValue: "2025-01-20",
            driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
            reference: "DS-1-1570",
          },
        ],
      }),
      createDriveClient: async () => ({}) as any,
      getRootFolderId: () => "root-folder-id",
      downloadDrivePdf: async () => Buffer.from("%PDF-1.4\nsource-drive"),
      createFolderPath: async () => ({ yearFolderId: "year", monthFolderId: "month" }),
      mergePdf: async () => Buffer.from("%PDF-1.4\nmerged"),
      uploadFile: async () => ({ id: "uploaded-id", url: "https://drive.google.com/file/d/uploaded-id/view" }),
    };

    const app = buildApp(deps);
    const response = await request(app).post("/causation/build").send({
      openaiFileIdRefs: [
        {
          name: "DS-1-1570.pdf",
          id: "file-123",
          mime_type: "application/pdf",
          download_link: "https://files.example.com/download/file-123",
        },
      ],
      debug: true,
    });

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(response.body.data.reference, "DS-1-1570");
    assert.equal(response.body.data.uploaded_file_name, "DS-1-1570.pdf");
  });

  it("rejects request without openaiFileIdRefs when document is missing", async () => {
    const deps = {
      readRegistroRows: async () => ({ spreadsheetId: "sheet-id", gid: "42421166", rows: [] }),
      createDriveClient: async () => ({}),
      getRootFolderId: () => "root",
      downloadDrivePdf: async () => Buffer.from(""),
      createFolderPath: async () => ({ yearFolderId: "year", monthFolderId: "month" }),
      mergePdf: async () => Buffer.from(""),
      uploadFile: async () => ({ id: "id", url: "url" }),
    };

    const app = buildApp(deps);
    const response = await request(app).post("/causation/build").send({ debug: true });

    assert.equal(response.status, 400);
    assert.equal(response.body.code, "missing_input_file");
  });

  it("rejects request with empty openaiFileIdRefs", async () => {
    const deps = {
      readRegistroRows: async () => ({ spreadsheetId: "sheet-id", gid: "42421166", rows: [] }),
      createDriveClient: async () => ({}),
      getRootFolderId: () => "root",
      downloadDrivePdf: async () => Buffer.from(""),
      createFolderPath: async () => ({ yearFolderId: "year", monthFolderId: "month" }),
      mergePdf: async () => Buffer.from(""),
      uploadFile: async () => ({ id: "id", url: "url" }),
    };

    const app = buildApp(deps);
    const response = await request(app).post("/causation/build").send({ openaiFileIdRefs: [] });

    assert.equal(response.status, 400);
    assert.equal(response.body.code, "missing_input_file");
    assert.equal(response.body.message, "openaiFileIdRefs debe contener al menos un archivo");
  });

  it("accepts openaiFileIdRefs nested inside params", async () => {
    global.fetch = async () =>
      new Response(Buffer.from("%PDF-1.4\nchatgpt-file"), {
        status: 200,
        headers: { "content-type": "application/pdf" },
      });

    const deps = {
      readRegistroRows: async () => ({
        spreadsheetId: "sheet-id",
        gid: "42421166",
        rows: [
          {
            rowNumber: 7,
            dateValue: "2025-01-20",
            driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
            reference: "DS-1-1570",
          },
        ],
      }),
      createDriveClient: async () => ({}) as any,
      getRootFolderId: () => "root-folder-id",
      downloadDrivePdf: async () => Buffer.from("%PDF-1.4\nsource-drive"),
      createFolderPath: async () => ({ yearFolderId: "year", monthFolderId: "month" }),
      mergePdf: async () => Buffer.from("%PDF-1.4\nmerged"),
      uploadFile: async () => ({ id: "uploaded-id", url: "https://drive.google.com/file/d/uploaded-id/view" }),
    };

    const app = buildApp(deps);
    const response = await request(app).post("/causation/build").send({
      params: {
        openaiFileIdRefs: [
          {
            name: "DS-1-1570.pdf",
            id: "file-123",
            mime_type: null,
            download_link: "https://files.example.com/download/file-123",
          },
        ],
      },
      debug: true,
    });

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(response.body.data.reference, "DS-1-1570");
  });

  it("rejects request when first openai file is not PDF", async () => {
    const deps = {
      readRegistroRows: async () => ({ spreadsheetId: "sheet-id", gid: "42421166", rows: [] }),
      createDriveClient: async () => ({}),
      getRootFolderId: () => "root",
      downloadDrivePdf: async () => Buffer.from(""),
      createFolderPath: async () => ({ yearFolderId: "year", monthFolderId: "month" }),
      mergePdf: async () => Buffer.from(""),
      uploadFile: async () => ({ id: "id", url: "url" }),
    };

    const app = buildApp(deps);
    const response = await request(app).post("/causation/build").send({
      openaiFileIdRefs: [
        {
          name: "archivo.txt",
          id: "file-123",
          mime_type: "text/plain",
          download_link: "https://files.example.com/download/file-123",
        },
      ],
    });

    assert.equal(response.status, 422);
    assert.equal(response.body.code, "unsupported_openai_file_mime_type");
  });

  it("rejects request when download_link is missing", async () => {
    const deps = {
      readRegistroRows: async () => ({ spreadsheetId: "sheet-id", gid: "42421166", rows: [] }),
      createDriveClient: async () => ({}),
      getRootFolderId: () => "root",
      downloadDrivePdf: async () => Buffer.from(""),
      createFolderPath: async () => ({ yearFolderId: "year", monthFolderId: "month" }),
      mergePdf: async () => Buffer.from(""),
      uploadFile: async () => ({ id: "id", url: "url" }),
    };

    const app = buildApp(deps);
    const response = await request(app).post("/causation/build").send({
      openaiFileIdRefs: [
        {
          name: "DS-1-1570.pdf",
          id: "file-123",
          mime_type: "application/pdf",
        },
      ],
    });

    assert.equal(response.status, 400);
    assert.equal(response.body.code, "missing_openai_download_link");
  });

  it("keeps multipart fallback with document file", async () => {
    const deps = {
      readRegistroRows: async () => ({
        spreadsheetId: "sheet-id",
        gid: "42421166",
        rows: [
          {
            rowNumber: 7,
            dateValue: "2025-01-20",
            driveLink: "https://drive.google.com/file/d/1AbcDefGhIjKlMnOpQrStUvWxYz12345/view",
            reference: "DS-1-1570",
          },
        ],
      }),
      createDriveClient: async () => ({}) as any,
      getRootFolderId: () => "root-folder-id",
      downloadDrivePdf: async () => Buffer.from("%PDF-1.4\nsource-drive"),
      createFolderPath: async () => ({ yearFolderId: "year", monthFolderId: "month" }),
      mergePdf: async () => Buffer.from("%PDF-1.4\nmerged"),
      uploadFile: async () => ({ id: "uploaded-id", url: "https://drive.google.com/file/d/uploaded-id/view" }),
    };

    const app = buildApp(deps);
    const response = await request(app)
      .post("/causation/build")
      .attach("document", Buffer.from("%PDF-1.4\nmanual-upload"), "DS-1-1570.pdf")
      .field("debug", "true");

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(response.body.data.debug.input_source, "multipart");
  });
});

const originalFetch = global.fetch;
