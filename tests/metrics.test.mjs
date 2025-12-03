import test from "node:test";
import assert from "node:assert/strict";
import { __testables } from "../payroll-recorder/src/metrics.js";

const { parseNumber, normalize, sumColumn, sumRawData, locateHeader } = __testables;

test("parseNumber handles parentheses and symbols", () => {
    assert.equal(parseNumber("(1,250.50)"), -1250.5);
    assert.equal(parseNumber("$2,000"), 2000);
    assert.equal(parseNumber(" "), 0);
});

test("normalize lowercases and trims values", () => {
    assert.equal(normalize(" PayRoll "), "payroll");
    assert.equal(normalize(null), "");
});

test("sumColumn totals matching header column", () => {
    const rows = [
        ["", ""],
        ["Amount", "Other"],
        [150, "x"],
        ["200.25", "y"],
        ["", ""]
    ];
    assert.equal(sumColumn(rows), 350.25);
});

test("locateHeader resolves first populated row", () => {
    const rows = [
        ["", ""],
        ["Employee", "Amount"]
    ];
    const { dataStartRow, indexMap } = locateHeader(rows);
    assert.equal(dataStartRow, 2);
    assert.equal(indexMap.employee, 0);
    assert.equal(indexMap.amount, 1);
});

test("sumRawData ignores totals and blank employees", () => {
    const rows = [
        ["", "", "", ""],
        ["Employee", "Department", "Gross Pay", "Bonus"],
        ["Alice", "Ops", "1000", 50],
        ["", "Ops", 200, 25],
        ["Totals", "", 9999, 9999]
    ];
    // Gross Pay column skipped, totals row ignored
    assert.equal(sumRawData(rows), 50);
});
