import { Type } from "typebox";

const loose = { additionalProperties: true } as const;
const NumOrNull = Type.Union([Type.Number(), Type.Null()]);

export const SectionSchema = Type.Object(
  { id: Type.String(), name: Type.String(), path: Type.String() },
  loose,
);
export const NotebookSchema = Type.Object(
  { id: Type.String(), name: Type.String(), path: Type.String(), sections: Type.Array(SectionSchema) },
  loose,
);
export const PageMetaSchema = Type.Object(
  { id: Type.String(), name: Type.String(), level: Type.String(), created: Type.String() },
  loose,
);

const PageLineSchema = Type.Object(
  {
    index: Type.Number(),
    objectID: Type.String(),
    text: Type.String(),
    inTitle: Type.Boolean(),
    inTable: Type.Boolean(),
    hasImage: Type.Boolean(),
    hasTable: Type.Boolean(),
    hasInk: Type.Boolean(),
    list: Type.String(),
    depth: Type.Number(),
    outlineID: Type.String(),
    quickStyle: Type.String(),
    alignment: Type.String(),
  },
  loose,
);
const PageOutlineSchema = Type.Object(
  {
    objectID: Type.String(),
    x: NumOrNull,
    y: NumOrNull,
    z: NumOrNull,
    width: NumOrNull,
    height: NumOrNull,
    inTitle: Type.Boolean(),
    lineCount: Type.Number(),
  },
  loose,
);
const PageImageSchema = Type.Object(
  {
    objectID: Type.String(),
    format: Type.String(),
    alt: Type.String(),
    x: NumOrNull,
    y: NumOrNull,
    width: NumOrNull,
    height: NumOrNull,
    outlineID: Type.String(),
    embedded: Type.Boolean(),
    callbackID: Type.String(),
  },
  loose,
);
const PageTableSchema = Type.Object(
  { objectID: Type.String(), rows: Type.Number(), columns: Type.Number(), outlineID: Type.String() },
  loose,
);

export const PageSchema = Type.Object(
  {
    id: Type.String(),
    name: Type.String(),
    level: Type.String(),
    lang: Type.String(),
    created: Type.String(),
    lastModified: Type.String(),
    lines: Type.Array(PageLineSchema),
    outlines: Type.Array(PageOutlineSchema),
    images: Type.Array(PageImageSchema),
    tables: Type.Array(PageTableSchema),
    rawXml: Type.Optional(Type.String()),
  },
  loose,
);

/** A dry run is a success too, so one shape covers both outcomes of a write. */
export const WriteResultSchema = Type.Object(
  {
    status: Type.Union([Type.Literal("dry-run"), Type.Literal("applied")]),
    applyToken: Type.Optional(Type.String()),
    preview: Type.Optional(Type.String()),
    page: Type.Optional(PageSchema),
    changes: Type.Optional(Type.Array(Type.Any())),
  },
  loose,
);

export const TARGET_PARAMS = {
  pageId: Type.Optional(Type.String({ description: "Page ID from pages or search. Preferred." })),
  notebook: Type.Optional(
    Type.String({ description: "Notebook name, used with section and page when pageId is unknown" }),
  ),
  section: Type.Optional(Type.String({ description: "Section name" })),
  page: Type.Optional(Type.String({ description: "Page name" })),
};

export const WRITE_PARAMS = {
  dryRun: Type.Optional(
    Type.Boolean({
      description: "Default true: preview only. Pass false, together with applyToken, to apply.",
    }),
  ),
  applyToken: Type.Optional(
    Type.String({
      description: "The applyToken from the dry run of this same call. Required when dryRun is false.",
    }),
  ),
};
