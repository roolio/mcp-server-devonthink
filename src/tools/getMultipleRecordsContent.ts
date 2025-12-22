import { z } from "zod";
import { createDevonThinkTool } from "./base/DevonThinkTool.js";

const GetMultipleRecordsContentSchema = z
	.object({
		uuids: z
			.array(z.string())
			.min(1)
			.max(100)
			.describe("Array of record UUIDs to retrieve content from (max 100)"),
		includeMetadata: z
			.boolean()
			.optional()
			.default(true)
			.describe("Include tags, dates, and location metadata"),
		databaseName: z
			.string()
			.optional()
			.describe("Optionally verify all records are from this database"),
	})
	.strict();

export const getMultipleRecordsContentTool = createDevonThinkTool({
	name: "get_multiple_records_content",
	description:
		'Retrieve content from multiple DEVONthink records in a single call. Optimized for batch operations and RAG workflows.\n\nExample:\n{\n  "uuids": ["uuid1", "uuid2", "uuid3"],\n  "includeMetadata": true\n}',

	inputSchema: GetMultipleRecordsContentSchema,

	buildScript: (input, helpers) => {
		const { uuids, includeMetadata, databaseName } = input;

		return helpers.wrapInTryCatch(`
      const theApp = Application("DEVONthink");
      theApp.includeStandardAdditions = true;

      // Check if DEVONthink is running
      if (!theApp.running()) {
        const result = {};
        result["success"] = false;
        result["error"] = "DEVONthink is not running";
        return JSON.stringify(result);
      }

      const uuids = ${helpers.formatValue(uuids)};
      const includeMetadata = ${includeMetadata};
      const targetDatabase = ${databaseName ? helpers.formatValue(databaseName) : "null"};

      const documents = [];
      const errors = [];

      for (const uuid of uuids) {
        try {
          const record = theApp.getRecordWithUuid(uuid);

          if (!record) {
            const err = {};
            err["uuid"] = uuid;
            err["error"] = "Record not found";
            errors.push(err);
            continue;
          }

          // Verify database if specified
          if (targetDatabase && record.database().name() !== targetDatabase) {
            const err = {};
            err["uuid"] = uuid;
            err["error"] = "Record not in database: " + targetDatabase;
            errors.push(err);
            continue;
          }

          // Build document with bracket notation (required for JXA)
          const doc = {};
          doc["uuid"] = uuid;
          doc["name"] = record.name();
          doc["content"] = record.plainText();

          if (includeMetadata) {
            doc["recordType"] = record.type();
            doc["tags"] = record.tags();
            doc["creationDate"] = record.creationDate() ? record.creationDate().toString() : null;
            doc["modificationDate"] = record.modificationDate() ? record.modificationDate().toString() : null;
            doc["location"] = record.location();
          }

          documents.push(doc);

        } catch (e) {
          const err = {};
          err["uuid"] = uuid;
          err["error"] = e.toString();
          errors.push(err);
        }
      }

      const result = {};
      result["success"] = true;
      result["documents"] = documents;
      result["errors"] = errors;
      result["totalRequested"] = uuids.length;
      result["totalRetrieved"] = documents.length;

      return JSON.stringify(result);
    `);
	},
});
