export function formatFourPartVersion(version) {
	if (!version || typeof version !== "string") {
		throw new Error("Invalid version string");
	}
	const clean = version.replace(/^v/, "").split("-")[0].split("+")[0].trim();
	const parts = clean.split(".");
	while (parts.length < 3) {
		parts.push("0");
	}
	return `${parts.slice(0, 3).join(".")}.0`;
}

export function updateManifestVersion(xmlContent, fourPartVersion) {
	if (!xmlContent.includes("<Version>")) {
		throw new Error("Manifest XML does not contain a <Version> element");
	}
	return xmlContent.replace(
		/<Version>[^<]*<\/Version>/,
		`<Version>${fourPartVersion}</Version>`,
	);
}
