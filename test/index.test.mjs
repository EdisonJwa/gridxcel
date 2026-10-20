
import * as g from "../src/index.js";
console.log("VERSION", g.VERSION);
console.log("exports:", Object.keys(g).sort().join(", "));
const s = g.normalizeSheet([["a", 1, true, new Date(2024,0,1)], ["b", 2]], "S");
console.log("sheet:", s.name, s.rowCount, s.columnCount, JSON.stringify(s.rows[0].map(c => [c.t, c.t === "date" ? c.v.toISOString().slice(0,10) : c.v])));
const b = g.normalizeBook({ SheetNames: ["One"], Sheets: { One: [["x", 1]] } });
console.log("book sheets:", b.sheets.length, b.sheets[0].name, b.sheets[0].rowCount);
console.log("address:", g.encodeCell(0,0), g.encodeCol(16383), JSON.stringify(g.decodeRange("B2:D5")));
console.log("numfmt:", g.formatNumber(1234.5678, "#,##0.00"), g.isDateFormat("yyyy-mm-dd"));
console.log("crc32('123456789'):", g.crc32(new TextEncoder().encode("123456789")).toString(16));
console.log("INDEX SMOKE OK");
