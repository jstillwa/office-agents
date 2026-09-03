import fs from "fs";
import path from "path";
import { describe, expect, it } from "vitest";

describe("manifest and security headers", () => {
  const root = path.resolve(__dirname, "../../..");

  it("declares ExcelApi 1.7 in excel manifest.xml and manifest.prod.xml", () => {
    const devManifest = fs.readFileSync(
      path.resolve(root, "packages/excel/manifest.xml"),
      "utf8",
    );
    const prodManifest = fs.readFileSync(
      path.resolve(root, "packages/excel/manifest.prod.xml"),
      "utf8",
    );

    expect(devManifest).toContain('<Set Name="ExcelApi" MinVersion="1.7"/>');
    expect(prodManifest).toContain('<Set Name="ExcelApi" MinVersion="1.7"/>');
  });

  it("ensures powerpoint dev and prod manifests have distinct GUIDs", () => {
    const devManifest = fs.readFileSync(
      path.resolve(root, "packages/powerpoint/manifest.xml"),
      "utf8",
    );
    const prodManifest = fs.readFileSync(
      path.resolve(root, "packages/powerpoint/manifest.prod.xml"),
      "utf8",
    );

    const devIdMatch = devManifest.match(/<Id>([^<]+)<\/Id>/);
    const prodIdMatch = prodManifest.match(/<Id>([^<]+)<\/Id>/);

    expect(devIdMatch).toBeTruthy();
    expect(prodIdMatch).toBeTruthy();
    expect(devIdMatch![1]).not.toBe(prodIdMatch![1]);
  });

  it("declares WebApplicationInfo in all prod manifests with access_as_user scope", () => {
    for (const app of ["excel", "powerpoint", "word"]) {
      const prodManifest = fs.readFileSync(
        path.resolve(root, `packages/${app}/manifest.prod.xml`),
        "utf8",
      );
      expect(prodManifest).toContain("<WebApplicationInfo");
      expect(prodManifest).toContain("<Scope>access_as_user</Scope>");
      expect(prodManifest).toContain("api://{HOST}/{CLIENT_ID}");
      expect(prodManifest).not.toContain("contoso.com");
    }
  });

  it("verifies public/assets icons exist for powerpoint and word", () => {
    const sizes = [16, 32, 64, 80];
    for (const app of ["powerpoint", "word"]) {
      for (const size of sizes) {
        const iconPath = path.resolve(
          root,
          `packages/${app}/public/assets/icon-${size}.png`,
        );
        expect(fs.existsSync(iconPath)).toBe(true);
      }
    }
  });

  it("verifies _headers contains required CSP directives and headers", () => {
    for (const app of ["excel", "powerpoint", "word"]) {
      const headersPath = path.resolve(
        root,
        `packages/${app}/public/_headers`,
      );
      const content = fs.readFileSync(headersPath, "utf8");
      expect(content).toContain("frame-ancestors 'self'");
      expect(content).toContain("https://*.office.com");
      expect(content).toContain("https://*.officeapps.live.com");
      expect(content).toContain("https://*.sharepoint.com");
      expect(content).toContain("https://*.office365.com");
      expect(content).toContain("default-src 'self'");
      expect(content).toContain(
        "script-src 'self' https://appsforoffice.microsoft.com",
      );
      expect(content).toContain("object-src 'none'");
      expect(content).toContain("X-Content-Type-Options: nosniff");
      expect(content).toContain(
        "Referrer-Policy: strict-origin-when-cross-origin",
      );
    }
  });
});
