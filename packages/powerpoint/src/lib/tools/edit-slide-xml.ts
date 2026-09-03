import type { AgentContext } from "@office-agents/core";
import { sandboxedEval } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import { escapeXml, parseXml } from "../pptx/xml-utils";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

export const MAX_SCRIPT_SIZE = 100 * 1024; // 100KB

export function validateSlideIndex(slideIndex: unknown): number {
  if (
    typeof slideIndex !== "number" ||
    !Number.isInteger(slideIndex) ||
    slideIndex < 0
  ) {
    throw new Error("slide_index must be a non-negative integer");
  }
  return slideIndex;
}

export function validateScriptCode(code: unknown): string {
  if (typeof code !== "string" || !code.trim()) {
    throw new Error("code must be a non-empty string");
  }
  if (code.length > MAX_SCRIPT_SIZE) {
    throw new Error(
      `Script size (${code.length} bytes) exceeds maximum limit of ${MAX_SCRIPT_SIZE} bytes`,
    );
  }
  return code;
}

export function createEditSlideXmlTool(ctx: AgentContext) {
  return defineTool({
    name: "edit_slide_xml",
    label: "Edit Slide XML",
    description:
      "Edit raw OOXML of a PowerPoint slide. Use for advanced formatting, " +
      "custom XML manipulation, diagrams, or anything not covered by other tools. " +
      "Use escapeXml(text) to escape special characters when embedding text in XML. " +
      "Use parseXml(xmlString) to parse XML safely.",
    parameters: Type.Object({
      slide_index: Type.Number({
        description:
          "0-based slide index (user's slide 1 = index 0, slide 3 = index 2)",
      }),
      code: Type.String({
        description:
          "Async function body receiving { zip, markDirty }. zip is a JSZip archive of the slide. " +
          "Call markDirty() if you modified files. " +
          "Globals: escapeXml(text) for safe XML text embedding, " +
          "parseXml(xmlString) for safe XML parsing (application/xml), " +
          "readFile(path) returns Promise<string> and readFileBuffer(path) returns Promise<Uint8Array> " +
          "to read files from the virtual filesystem (e.g. uploaded images, SVGs). " +
          "writeFile(path, content) returns Promise<void> to write string or Uint8Array to the virtual filesystem.",
      }),
      explanation: Type.Optional(
        Type.String({
          description: "Brief description (max 50 chars)",
          maxLength: 50,
        }),
      ),
    }),
    execute: async (_toolCallId, params) => {
      try {
        validateSlideIndex(params.slide_index);
        validateScriptCode(params.code);

        const result = await safeRun(async (context) => {
          return withSlideZip(context, params.slide_index, async (args) => {
            return sandboxedEval(params.code, {
              ...args,
              escapeXml,
              parseXml,
              readFile: (path: string) => ctx.readFile(path),
              readFileBuffer: (path: string) => ctx.readFileBuffer(path),
              writeFile: (path: string, content: string | Uint8Array) =>
                ctx.writeFile(path, content),
              XMLSerializer,
            });
          });
        });

        return toolSuccess({
          success: true,
          result: result !== undefined ? result : null,
        });
      } catch (error) {
        const message =
          error instanceof Error ? error.message : "Failed to edit slide XML";
        return toolError(message);
      }
    },
    modifiedSlide: (params) => params.slide_index,
  });
}
