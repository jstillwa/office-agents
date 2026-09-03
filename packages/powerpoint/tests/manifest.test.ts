import fs from "fs";
import path from "path";
import { describe, expect, it } from "vitest";

describe("powerpoint manifest and assets", () => {
  const pkgDir = path.resolve(__dirname, "..");

  it("has distinct GUID in dev manifest vs prod manifest", () => {
    const devManifest = fs.readFileSync(
      path.resolve(pkgDir, "manifest.xml"),
      "utf8",
    );
    const prodManifest = fs.readFileSync(
      path.resolve(pkgDir, "manifest.prod.xml"),
      "utf8",
    );

    const devId = devManifest.match(/<Id>([^<]+)<\/Id>/)?.[1];
    const prodId = prodManifest.match(/<Id>([^<]+)<\/Id>/)?.[1];

    expect(devId).toBeDefined();
    expect(prodId).toBeDefined();
    expect(devId).not.toBe(prodId);
  });

  it("declares WebApplicationInfo in prod manifest", () => {
    const prodManifest = fs.readFileSync(
      path.resolve(pkgDir, "manifest.prod.xml"),
      "utf8",
    );
    expect(prodManifest).toContain("<WebApplicationInfo");
    expect(prodManifest).toContain("<Scope>access_as_user</Scope>");
    expect(prodManifest).not.toContain("contoso.com");
  });

  it("has icon assets in public/assets", () => {
    for (const size of [16, 32, 64, 80]) {
      expect(
        fs.existsSync(path.resolve(pkgDir, `public/assets/icon-${size}.png`)),
      ).toBe(true);
    }
  });

  it("has frame-ancestors in public/_headers", () => {
    const headers = fs.readFileSync(
      path.resolve(pkgDir, "public/_headers"),
      "utf8",
    );
    expect(headers).toContain("frame-ancestors 'self'");
  });
});
