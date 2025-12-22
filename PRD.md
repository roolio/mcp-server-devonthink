# PRD: Add Batch Content Retrieval Tool to DEVONthink MCP Server

## Context

Currently, the MCP server only supports retrieving content from a single DEVONthink record at a time via `get_record_content`. For workflows that need to analyze multiple documents (e.g., RAG-style queries across 10-50 notes), this requires N sequential tool calls, resulting in poor performance:

- **Current**: N × (5-20ms) = 50-200ms for 10 documents
- **Target**: 1 × 30ms for 10 documents (5x improvement)

The bottleneck is JSON-RPC round-trips, not DEVONthink's API.

## Solution

Create a new tool `get_multiple_records_content` that retrieves content from multiple records in a single JXA execution, eliminating round-trip overhead.

## Technical Implementation

### 1. Create New Tool File

**Location**: `src/tools/getMultipleRecordsContent.ts`

**Schema**:
```typescript
const GetMultipleRecordsContentSchema = z.object({
  uuids: z.array(z.string()).min(1).max(100).describe("Array of record UUIDs to retrieve content from (max 100)"),
  includeMetadata: z.boolean().optional().default(true).describe("Include tags, dates, and location metadata"),
  databaseName: z.string().optional().describe("Optionally verify all records are from this database")
}).strict();
```

**Return Type**:
```typescript
interface GetMultipleRecordsContentResult {
  success: boolean;
  error?: string;
  documents?: Array<{
    uuid: string;
    name: string;
    content: string;
    recordType?: string;
    tags?: string[];
    creationDate?: string;
    modificationDate?: string;
    location?: string;
  }>;
  errors?: Array<{
    uuid: string;
    error: string;
  }>;
  totalRequested: number;
  totalRetrieved: number;
}
```

### 2. JXA Script Implementation

**Key Points**:
- Loop through UUIDs inside JXA (not in TypeScript)
- Use `record.plainText()` for content (handles markdown, txt, formatted notes)
- Handle errors per-document (don't fail entire batch)
- Build result objects with bracket notation (JXA limitation)
- Return both successful retrievals and errors

**Script Pattern**:
```javascript
(() => {
  const theApp = Application("DEVONthink");
  theApp.includeStandardAdditions = true;
  
  const uuids = ["uuid1", "uuid2", "uuid3"];
  const includeMetadata = true;
  const targetDatabase = "roolio"; // or null
  
  const results = [];
  const errors = [];
  
  for (const uuid of uuids) {
    try {
      const record = theApp.getRecordWithUuid(uuid);
      
      if (!record) {
        errors.push({ uuid: uuid, error: "Record not found" });
        continue;
      }
      
      // Optional: verify database
      if (targetDatabase && record.database().name() !== targetDatabase) {
        errors.push({ 
          uuid: uuid, 
          error: "Record not in database: " + targetDatabase 
        });
        continue;
      }
      
      // Build result using bracket notation (REQUIRED for JXA)
      const result = {};
      result["uuid"] = uuid;
      result["name"] = record.name();
      result["content"] = record.plainText();
      
      if (includeMetadata) {
        result["recordType"] = record.recordType();
        result["tags"] = record.tags();
        result["creationDate"] = record.creationDate() ? record.creationDate().toString() : null;
        result["modificationDate"] = record.modificationDate() ? record.modificationDate().toString() : null;
        result["location"] = record.location();
      }
      
      results.push(result);
      
    } catch (e) {
      errors.push({ uuid: uuid, error: e.toString() });
    }
  }
  
  const response = {};
  response["success"] = true;
  response["documents"] = results;
  response["errors"] = errors;
  response["totalRequested"] = uuids.length;
  response["totalRetrieved"] = results.length;
  
  return JSON.stringify(response);
})();
```

### 3. Integration

**Add to `src/devonthink.ts`**:
```typescript
import { getMultipleRecordsContentTool } from "./tools/getMultipleRecordsContent.js";

const tools: Tool[] = [
  // ... existing tools
  getMultipleRecordsContentTool,
];
```

### 4. Update Documentation

**Add to `CLAUDE.md`**:

In the "Available Tools" section:
```markdown
28. **`get_multiple_records_content`** - Retrieve content from multiple records in a single call (batch operation, optimized for RAG workflows)
```

In the "Project Structure" section:
```markdown
- **`getMultipleRecordsContent.ts`**: Retrieves content from multiple records efficiently
```

## Critical JXA Constraints

⚠️ **MUST follow these rules** (see CLAUDE.md "JXA Interpreter Limitations"):

1. **Object building**: Use bracket notation ONLY
```javascript
   // ❌ WRONG - will fail
   const obj = { uuid: uuid, name: name };
   
   // ✅ CORRECT
   const obj = {};
   obj["uuid"] = uuid;
   obj["name"] = name;
```

2. **No console.log**: Causes MCP errors (outputs to stderr)

3. **String interpolation**: Use proper escaping
```javascript
   const uuid = ${formatValueForJXA(input.uuids)};
```

4. **Error handling**: Always wrap in try-catch per document

## Testing Checklist

After implementation, test these scenarios:

### Basic Tests
- [ ] Retrieve 1 document
- [ ] Retrieve 10 documents
- [ ] Retrieve 50 documents
- [ ] Empty array (should error with validation)

### Error Handling
- [ ] One invalid UUID in batch (should continue with others)
- [ ] All invalid UUIDs (should return empty documents array)
- [ ] Record exists but in different database (with databaseName filter)

### Metadata Tests
- [ ] includeMetadata: true (should include tags, dates, location)
- [ ] includeMetadata: false (should only include uuid, name, content)

### Performance
- [ ] Compare timing: 10× get_record_content vs 1× get_multiple_records_content
- [ ] Expected: 3-5x speedup

## Build & Test Commands
```bash
# Build
npm run build

# Format check (REQUIRED before commit)
npm run format:check

# Auto-format
npm run format

# Run tests
npm test
```

## Success Criteria

1. ✅ Tool successfully retrieves content from multiple records
2. ✅ Errors in individual records don't fail entire batch
3. ✅ Performance is 3-5x faster than sequential calls
4. ✅ All existing tests still pass
5. ✅ Code passes format check (`npm run format:check`)
6. ✅ Tool is documented in CLAUDE.md

## Reference Files

- **Example tool with batch pattern**: `src/tools/ai/askAiAboutDocuments.ts`
- **Example tool with metadata**: `src/tools/getRecordContent.ts`
- **JXA limitations**: See `CLAUDE.md` section "JXA Interpreter Limitations"
- **Validation patterns**: Other tools in `src/tools/`

## Notes for Implementation

1. Start by copying structure from `getRecordContent.ts`
2. Modify to accept array of UUIDs instead of single UUID
3. Move the loop INSIDE the JXA script (critical for performance)
4. Use bracket notation for all object building
5. Test incrementally (1 doc → 5 docs → 50 docs)
6. Run `npm run format` before committing
