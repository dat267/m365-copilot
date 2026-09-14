import test from "node:test";
import assert from "node:assert/strict";

import { isImage, planImageBatches, planAttachmentBatches, renderAttachmentManifest } from "./attachments.js";

test("isImage trusts an image content type", () => {
  assert.equal(isImage({ name: "x.bin", contentType: "image/png" }), true);
});

test("isImage falls back to the file extension when the type is missing", () => {
  assert.equal(isImage({ name: "Screenshot.PNG", contentType: "" }), true);
});

test("isImage rejects a non-image", () => {
  assert.equal(isImage({ name: "logs.txt", contentType: "text/plain" }), false);
});

const img = (n) => ({ name: `shot-${n}.png`, contentType: "image/png" });

test("planImageBatches splits images into batches that respect the message cap", () => {
  const plan = planImageBatches([img(1), img(2), img(3), img(4), img(5), img(6), img(7)], {
    maxPerMessage: 3,
  });

  assert.deepEqual(plan.batches.map((b) => b.length), [3, 3, 1]);
  assert.equal(plan.images.length, 7);
});

test("planImageBatches keeps non-images out of the image batches", () => {
  const doc = { name: "trace.log", contentType: "text/plain" };

  const plan = planImageBatches([img(1), doc], { maxPerMessage: 3 });

  assert.deepEqual(plan.batches.map((b) => b.length), [1]);
  assert.deepEqual(plan.files.map((f) => f.name), ["trace.log"]);
});

test("planImageBatches returns no batches when there are no images", () => {
  const plan = planImageBatches([{ name: "a.pdf", contentType: "application/pdf" }]);

  assert.deepEqual(plan.batches, []);
  assert.equal(plan.files.length, 1);
});

test("renderAttachmentManifest lists name, type, size and source", () => {
  const md = renderAttachmentManifest([
    {
      name: "Issue_Query.png",
      contentType: "image/png",
      size: 22065,
      source: "conversation 43940693585",
      url: "https://myzoi.attachments.freshservice.com/x/Issue_Query.png",
    },
  ]);

  assert.match(md, /## Attachments/);
  assert.match(md, /Issue_Query\.png/);
  assert.match(md, /image\/png/);
  assert.match(md, /22065/);
  assert.match(md, /conversation 43940693585/);
});

test("renderAttachmentManifest states plainly when there are none", () => {
  const md = renderAttachmentManifest([]);

  assert.match(md, /## Attachments/);
  assert.match(md, /\(none\)/);
});

test("planAttachmentBatches shares one cap across images and files", () => {
  const items = [img(1), img(2), { name: "a.txt", contentType: "text/plain" }, { name: "b.pdf", contentType: "application/pdf" }, img(3)];

  const batches = planAttachmentBatches(items, { maxPerMessage: 3 });

  assert.deepEqual(batches.map((b) => b.length), [3, 2]);
  assert.ok(batches.every((b) => b.length <= 3), "no batch exceeds the shared cap");
  assert.equal(batches.flat().length, items.length, "nothing dropped");
});

test("planAttachmentBatches never puts 3 images and a file in one message", () => {
  const items = [img(1), img(2), img(3), { name: "a.txt", contentType: "text/plain" }];

  const batches = planAttachmentBatches(items);

  assert.deepEqual(batches.map((b) => b.length), [3, 1], "the file gets its own message");
});
