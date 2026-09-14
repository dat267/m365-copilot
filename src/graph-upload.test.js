import test from "node:test";
import assert from "node:assert/strict";

import { spoId, toFileAnnotations } from "./graph-upload.js";

// Synthetic vector (a made-up `b!` drive id built from three GUIDs), so no real
// tenant/sharepoint identifiers live in the tests. Produced with the same
// mixed-endian GUID packing SharePoint uses; the algorithm is what's under test.
const DRIVE_ID = "b!ERERESIiMzNERFVVVVVVVaqqqqq7u8zM3d3u7u7u7u6ZmZmZiIh3d2ZmVVVVVVVV";
const ITEM_ID = "01SYNTHETICITEMID0000000000000000";
const EXPECTED =
  "SPO_MTExMTExMTEtMjIyMi0zMzMzLTQ0NDQtNTU1NTU1NTU1NTU1LGFhYWFhYWFhLWJiYmItY2NjYy1kZGRkLWVlZWVlZWVlZWVlZSw5OTk5OTk5OS04ODg4LTc3NzctNjY2Ni01NTU1NTU1NTU1NTU_01SYNTHETICITEMID0000000000000000";

test("spoId reproduces the id the web client sends", () => {
  assert.equal(spoId(DRIVE_ID, ITEM_ID), EXPECTED);
});

test("toFileAnnotations builds the LocalFile annotation shape", () => {
  const [a] = toFileAnnotations([
    { driveId: DRIVE_ID, itemId: ITEM_ID, fileName: "probe-file.txt", webUrl: "https://sp/x.txt" },
  ]);

  assert.deepEqual(a, {
    id: EXPECTED,
    text: "probe-file.txt",
    url: "https://sp/x.txt",
    messageAnnotationType: "LocalFile",
  });
});

test("toFileAnnotations skips entries without ids", () => {
  assert.deepEqual(toFileAnnotations([{ fileName: "x.txt" }]), []);
  assert.deepEqual(toFileAnnotations(), []);
});
