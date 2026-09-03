import { describe, expect, it } from "vitest";
import { formatFourPartVersion, updateManifestVersion } from "../../../scripts/release-utils.mjs";

describe("release-utils", () => {
	describe("formatFourPartVersion", () => {
		it("formats 3-part semver to 4-part Office version", () => {
			expect(formatFourPartVersion("1.2.3")).toBe("1.2.3.0");
			expect(formatFourPartVersion("0.2.10")).toBe("0.2.10.0");
		});

		it("handles versions with leading 'v'", () => {
			expect(formatFourPartVersion("v1.0.4")).toBe("1.0.4.0");
		});

		it("strips prerelease and build metadata", () => {
			expect(formatFourPartVersion("1.2.3-beta.1")).toBe("1.2.3.0");
			expect(formatFourPartVersion("2.0.0-rc.2+sha.12345")).toBe("2.0.0.0");
		});

		it("pads shorter semver to 3 parts before adding trailing 0", () => {
			expect(formatFourPartVersion("1.0")).toBe("1.0.0.0");
			expect(formatFourPartVersion("2")).toBe("2.0.0.0");
		});

		it("throws on invalid version input", () => {
			expect(() => formatFourPartVersion("")).toThrow();
			expect(() => formatFourPartVersion(null as unknown as string)).toThrow();
		});
	});

	describe("updateManifestVersion", () => {
		it("replaces <Version> element while preserving <VersionOverrides>", () => {
			const xml = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  <Id>d36b6541-f3fc-4c5c-9136-123ee8d0c71f</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts><Host xsi:type="Workbook" /></Hosts>
  </VersionOverrides>
</OfficeApp>`;

			const updated = updateManifestVersion(xml, "0.2.11.0");
			expect(updated).toContain("<Version>0.2.11.0</Version>");
			expect(updated).not.toContain("<Version>1.0.0.0</Version>");
			expect(updated).toContain('<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"');
		});

		it("throws when <Version> is missing", () => {
			const xml = `<OfficeApp><Id>123</Id></OfficeApp>`;
			expect(() => updateManifestVersion(xml, "1.0.0.0")).toThrow(
				"Manifest XML does not contain a <Version> element",
			);
		});
	});
});
