import { z } from "zod";
import { zodToJsonSchema } from "zod-to-json-schema";
import { Tool, ToolSchema } from "@modelcontextprotocol/sdk/types.js";
import { executeJxa } from "../applescript/execute.js";
import { escapeStringForJXA } from "../utils/escapeString.js";
import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";

const ToolInputSchema = ToolSchema.shape.inputSchema;
type ToolInput = z.infer<typeof ToolInputSchema>;

const CreateRecordSchema = z
	.object({
		name: z.string().describe("Name of the new record"),
		type: z
			.string()
			.describe("Record type (e.g., 'markdown', 'formatted note', 'bookmark', 'group')"),
		content: z.string().optional().describe("Content for text-based records (optional)"),
		url: z.string().optional().describe("URL for bookmark records (optional)"),
		parentGroupUuid: z
			.string()
			.optional()
			.describe("UUID of the parent group (optional, defaults to incoming group)"),
		databaseName: z
			.string()
			.optional()
			.describe("Database to create the record in (optional, defaults to current)"),
	})
	.strict();

type CreateRecordInput = z.infer<typeof CreateRecordSchema>;

/**
 * Sanitize a filename for the filesystem.
 * Removes or replaces characters that are problematic in file paths.
 */
function sanitizeFilename(name: string): string {
	return name
		.replace(/[/\\:*?"<>|]/g, "-") // Replace filesystem-unsafe chars
		.replace(/\s+/g, " ") // Normalize whitespace
		.trim()
		.slice(0, 200); // Limit length for filesystem compatibility
}

/**
 * Get file extension for a DEVONthink record type.
 */
function getExtensionForType(type: string): string | null {
	const extensionMap: Record<string, string> = {
		markdown: ".md",
		txt: ".txt",
		rtf: ".rtf",
		html: ".html",
	};
	return extensionMap[type.toLowerCase()] || null;
}

const createRecord = async (
	input: CreateRecordInput,
): Promise<{
	success: boolean;
	recordId?: number;
	name?: string;
	uuid?: string;
	error?: string;
}> => {
	const { name, type, content, url, parentGroupUuid, databaseName } = input;

	// Determine if we can use the temporary file approach
	const extension = getExtensionForType(type);
	const useFileImport = content && extension;

	if (useFileImport) {
		// Use sanitized name for the temporary file
		// This way DEVONthink will import with the correct name directly
		const safeFilename = sanitizeFilename(name) + extension;
		const tempFilePath = join(tmpdir(), safeFilename);
		// Escape the path for JXA (in case tmpdir contains special characters)
		const escapedTempFilePath = escapeStringForJXA(tempFilePath);
		const escapedName = escapeStringForJXA(name);

		try {
			// Write the content to a temporary file
			writeFileSync(tempFilePath, content, "utf-8");

			const script = `
				(() => {
					const theApp = Application("DEVONthink");
					theApp.includeStandardAdditions = true;

					try {
						let targetDatabase;
						if ("${databaseName || ""}") {
							const databases = theApp.databases();
							targetDatabase = databases.find(db => db.name() === "${escapeStringForJXA(databaseName || "")}");
							if (!targetDatabase) {
								throw new Error("Database not found: ${escapeStringForJXA(databaseName || "")}");
							}
						} else {
							targetDatabase = theApp.currentDatabase();
						}

						let destinationGroup;
						if ("${parentGroupUuid || ""}") {
							destinationGroup = theApp.getRecordWithUuid("${escapeStringForJXA(parentGroupUuid || "")}");
							if (!destinationGroup) {
								throw new Error("Parent group with UUID not found: ${escapeStringForJXA(parentGroupUuid || "")}");
							}
						} else {
							destinationGroup = targetDatabase.incomingGroup();
						}

						// Import the file
						const importedRecords = theApp.import("${escapedTempFilePath}", { to: destinationGroup });

						if (importedRecords && importedRecords.length > 0) {
							const newRecord = importedRecords[0];

							// Rename if the imported name doesn't match the desired name exactly
							// (can happen if the name contained non-filesystem-safe characters)
							const desiredName = "${escapedName}";
							if (desiredName && newRecord.name() !== desiredName) {
								newRecord.name = desiredName;
							}

							const result = {};
							result["success"] = true;
							result["recordId"] = newRecord.id();
							result["name"] = newRecord.name();
							result["uuid"] = newRecord.uuid();
							return JSON.stringify(result);
						} else {
							const result = {};
							result["success"] = false;
							result["error"] = "Failed to import file";
							return JSON.stringify(result);
						}
					} catch (error) {
						const result = {};
						result["success"] = false;
						result["error"] = error.toString();
						return JSON.stringify(result);
					}
				})();
			`;

			const result = await executeJxa<{
				success: boolean;
				recordId?: number;
				name?: string;
				uuid?: string;
				error?: string;
			}>(script);

			// Clean up the temporary file
			try {
				unlinkSync(tempFilePath);
			} catch {
				// Ignore cleanup errors
			}

			return result;
		} catch (error) {
			// Clean up on error
			try {
				unlinkSync(tempFilePath);
			} catch {
				// Ignore
			}
			return {
				success: false,
				error: error instanceof Error ? error.message : String(error),
			};
		}
	}

	// Fallback: use the legacy method (for bookmarks, groups, formatted notes, etc.)
	const script = `
		(() => {
			const theApp = Application("DEVONthink");
			theApp.includeStandardAdditions = true;

			try {
				let targetDatabase;
				if ("${databaseName || ""}") {
					const databases = theApp.databases();
					targetDatabase = databases.find(db => db.name() === "${escapeStringForJXA(databaseName || "")}");
					if (!targetDatabase) {
						throw new Error("Database not found: ${escapeStringForJXA(databaseName || "")}");
					}
				} else {
					targetDatabase = theApp.currentDatabase();
				}

				let destinationGroup;
				if ("${parentGroupUuid || ""}") {
					destinationGroup = theApp.getRecordWithUuid("${escapeStringForJXA(parentGroupUuid || "")}");
					if (!destinationGroup) {
						throw new Error("Parent group with UUID not found: ${escapeStringForJXA(parentGroupUuid || "")}");
					}
				} else {
					destinationGroup = targetDatabase.incomingGroup();
				}

				const recordProps = {};
				recordProps["name"] = "${escapeStringForJXA(name)}";
				recordProps["type"] = "${escapeStringForJXA(type)}";

				${url ? `recordProps["URL"] = "${escapeStringForJXA(url)}";` : ""}

				const newRecord = theApp.createRecordWith(recordProps, { in: destinationGroup });

				if (newRecord) {
					const result = {};
					result["success"] = true;
					result["recordId"] = newRecord.id();
					result["name"] = newRecord.name();
					result["uuid"] = newRecord.uuid();
					return JSON.stringify(result);
				} else {
					const result = {};
					result["success"] = false;
					result["error"] = "Failed to create record";
					return JSON.stringify(result);
				}
			} catch (error) {
				const result = {};
				result["success"] = false;
				result["error"] = error.toString();
				return JSON.stringify(result);
			}
		})();
	`;

	return await executeJxa<{
		success: boolean;
		recordId?: number;
		name?: string;
		uuid?: string;
		error?: string;
	}>(script);
};

export const createRecordTool: Tool = {
	name: "create_record",
	description:
		'Create a new record in DEVONthink.\n\nExample:\n{\n  "name": "New Note",\n  "type": "markdown",\n  "content": "# Hello World"\n}',
	inputSchema: zodToJsonSchema(CreateRecordSchema) as ToolInput,
	run: createRecord,
};
