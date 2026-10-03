/**
 * Vitest runs in Node, but the app imports `read-excel-file`'s browser entry
 * (the same code path the real app uses). That entry expects a global
 * `DOMParser`, which browsers provide natively. Supplying @xmldom/xmldom here
 * keeps the tests on the exact production code path instead of swapping in a
 * different module for tests.
 */
import { DOMParser } from "@xmldom/xmldom";

if (typeof globalThis.DOMParser === "undefined") {
  (globalThis as { DOMParser?: typeof DOMParser }).DOMParser = DOMParser;
}
