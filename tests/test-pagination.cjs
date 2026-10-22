// Run against worship-data.ts compiled to CommonJS in a temporary directory.
const assert = require("node:assert/strict");
const { parseWorshipData, buildPreview } = require(process.argv[2]);

function preview(song) {
  const scripture = { title: "Scripture", lines: ["Verse"] };
  return buildPreview(parseWorshipData(JSON.stringify({
    call_to_worship: scripture, hymns: [song], theme_scripture: scripture,
    response_hymn: song,
  }))).filter((slide) => slide.kind === "hymn" || slide.kind === "response_hymn")
    .map((slide) => slide.lines);
}

for (const marker of ["间奏", "（间奏）", "( 间奏 )", "结束", "（尾声）", "前奏"]) {
  for (const lyrics of [["A", "B"], ["A", "B", "C", "D"]]) {
    assert.deepEqual(preview({ title: "Test", lines: [...lyrics, "", marker, "E"] }),
      [[...lyrics, marker], ["E"], [...lyrics, marker], ["E"]]);
  }
}
assert.deepEqual(preview({ title: "Test", pages: [
  { lines: ["A", "B"], footer: "（间奏）" }, { lines: ["C"] },
  { lines: ["D", "E"], footer: "结束" },
]}), [["A", "B", "（间奏）"], ["C"], ["D", "E", "结束"],
     ["A", "B", "（间奏）"], ["C"], ["D", "E", "结束"]]);
assert.throws(() => preview({ title: "Test", lines: ["间奏", "A"] }));
for (const pages of [null, [], [{}], [{ lines: [] }], [{ lines: [" "] }],
  [{ lines: ["A", "间奏"] }], [{ lines: ["A"], footer: 3 }],
  [{ lines: ["A"], footer: "unknown" }]]) {
  assert.throws(() => preview({ title: "Test", pages }));
}
assert.throws(() => preview({ title: "Test", lines: ["different"], pages: [{ lines: ["A"] }] }));
console.log("Frontend lyric pagination checks passed.");
