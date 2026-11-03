import { defineCollection } from "astro:content";
import { glob } from "astro/loaders";
import { defineDocSchema, partialsCollection, withNimbusMarkdown } from "@cloudflare/nimbus-docs/content";
export const collections = {
  partials: defineCollection(partialsCollection()),
  docs: defineCollection({
    loader: withNimbusMarkdown(glob({ base: "../content", pattern: "**/*.md" })),
    schema: defineDocSchema(),
  }),
};
