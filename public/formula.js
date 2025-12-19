// Minimal Excel-like formula evaluator for this template.
// Supports: IF, SUM, TRIM, arithmetic + - * /, comparisons, &, cell refs (A1, $D$2), ranges (D4:D5).

function isNumberLike(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === "number") return Number.isFinite(v);
  if (typeof v === "string" && v.trim() !== "") return !Number.isNaN(Number(v));
  return false;
}

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const s = v.trim();
    if (s === "") return 0;
    const n = Number(s);
    return Number.isNaN(n) ? 0 : n;
  }
  return 0;
}

function toText(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

function truthyExcel(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (typeof v === "string") return v !== "";
  return Boolean(v);
}

function trimExcel(v) {
  // Excel TRIM also collapses internal multiple spaces; we mimic that.
  return toText(v).trim().replace(/\s+/g, " ");
}

function colToNum(col) {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function numToCol(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function normalizeRef(ref) {
  return ref.replace(/\$/g, "").toUpperCase();
}

function splitRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) throw new Error(`Bad cell ref: ${ref}`);
  return { col: m[1], row: Number(m[2]) };
}

function expandRange(rangeRef) {
  const [a, b] = rangeRef.split(":").map(normalizeRef);
  const ra = splitRef(a);
  const rb = splitRef(b);
  const colStart = Math.min(colToNum(ra.col), colToNum(rb.col));
  const colEnd = Math.max(colToNum(ra.col), colToNum(rb.col));
  const rowStart = Math.min(ra.row, rb.row);
  const rowEnd = Math.max(ra.row, rb.row);
  const out = [];
  for (let r = rowStart; r <= rowEnd; r++) {
    for (let c = colStart; c <= colEnd; c++) {
      out.push(`${numToCol(c)}${r}`);
    }
  }
  return out;
}

// --- Tokenizer / Parser ------------------------------------------------------

function tokenize(formula) {
  const s = formula.trim();
  const tokens = [];
  let i = 0;

  const push = (type, value) => tokens.push({ type, value });

  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // string literal "..."
    if (ch === '"') {
      let j = i + 1;
      let out = "";
      while (j < s.length) {
        if (s[j] === '"' && s[j + 1] === '"') {
          out += '"';
          j += 2;
          continue;
        }
        if (s[j] === '"') break;
        out += s[j];
        j++;
      }
      push("string", out);
      i = j + 1;
      continue;
    }

    // operators (2-char)
    const two = s.slice(i, i + 2);
    if (["<>", "<=", ">="].includes(two)) {
      push("op", two);
      i += 2;
      continue;
    }

    // single-char operators / punctuation
    if ("+-*/&(),:<>=".includes(ch)) {
      push("op", ch);
      i += 1;
      continue;
    }

    // number
    if (/\d/.test(ch) || (ch === "." && /\d/.test(s[i + 1] || ""))) {
      let j = i + 1;
      while (j < s.length && /[\d.]/.test(s[j])) j++;
      push("number", s.slice(i, j));
      i = j;
      continue;
    }

    // identifier or cell ref (may include $)
    if (/[A-Za-z_$]/.test(ch)) {
      let j = i + 1;
      while (j < s.length && /[A-Za-z0-9_$]/.test(s[j])) j++;
      const ident = s.slice(i, j);
      push("ident", ident);
      i = j;
      continue;
    }

    throw new Error(`Unexpected char '${ch}' in formula: ${formula}`);
  }

  return tokens;
}

function parse(formula) {
  const tokens = tokenize(formula);
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  const acceptOp = (op) => peek() && peek().type === "op" && peek().value === op && (next(), true);
  const expectOp = (op) => {
    if (!acceptOp(op)) throw new Error(`Expected '${op}' in formula: ${formula}`);
  };

  // Grammar:
  // expr -> compare
  // compare -> concat ( ( = | <> | < | > | <= | >= ) concat )*
  // concat -> add ( & add )*
  // add -> mul ( (+|-) mul )*
  // mul -> unary ( (*|/) unary )*
  // unary -> (+|-) unary | primary
  // primary -> number | string | cell | func | ( expr )

  function parseExpr() {
    return parseCompare();
  }

  function parseCompare() {
    let node = parseConcat();
    while (peek() && peek().type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(peek().value)) {
      const op = next().value;
      const right = parseConcat();
      node = { type: "bin", op, left: node, right };
    }
    return node;
  }

  function parseConcat() {
    let node = parseAdd();
    while (acceptOp("&")) {
      const right = parseAdd();
      node = { type: "bin", op: "&", left: node, right };
    }
    return node;
  }

  function parseAdd() {
    let node = parseMul();
    while (peek() && peek().type === "op" && ["+", "-"].includes(peek().value)) {
      const op = next().value;
      const right = parseMul();
      node = { type: "bin", op, left: node, right };
    }
    return node;
  }

  function parseMul() {
    let node = parseUnary();
    while (peek() && peek().type === "op" && ["*", "/"].includes(peek().value)) {
      const op = next().value;
      const right = parseUnary();
      node = { type: "bin", op, left: node, right };
    }
    return node;
  }

  function parseUnary() {
    if (peek() && peek().type === "op" && ["+", "-"].includes(peek().value)) {
      const op = next().value;
      const expr = parseUnary();
      return { type: "unary", op, expr };
    }
    return parsePrimary();
  }

  function parsePrimary() {
    const t = peek();
    if (!t) throw new Error(`Unexpected end of formula: ${formula}`);

    if (acceptOp("(")) {
      const e = parseExpr();
      expectOp(")");
      return e;
    }

    if (t.type === "number") {
      next();
      return { type: "num", value: Number(t.value) };
    }

    if (t.type === "string") {
      next();
      return { type: "str", value: t.value };
    }

    if (t.type === "ident") {
      const raw = next().value;
      const ident = raw.toUpperCase();

      // Function call?
      if (acceptOp("(")) {
        const args = [];
        if (!acceptOp(")")) {
          while (true) {
            args.push(parseExpr());
            if (acceptOp(")")) break;
            expectOp(",");
          }
        }
        return { type: "call", name: ident, args };
      }

      // Range (A1:D5) uses ':' operator which tokenizes separately; handle as node if next is ':'
      if (peek() && peek().type === "op" && peek().value === ":") {
        next(); // ':'
        const t2 = next();
        if (!t2 || t2.type !== "ident") throw new Error(`Bad range in formula: ${formula}`);
        return { type: "range", ref: `${raw}:${t2.value}` };
      }

      // Cell ref?
      if (/^\$?[A-Z]{1,3}\$?\d+$/.test(ident)) {
        return { type: "cell", ref: raw };
      }

      // Bare identifier as text (rare in this template)
      return { type: "name", name: ident };
    }

    throw new Error(`Unexpected token in formula: ${JSON.stringify(t)}`);
  }

  const ast = parseExpr();
  if (pos !== tokens.length) {
    throw new Error(`Could not fully parse formula: ${formula}`);
  }
  return ast;
}

// --- Evaluation --------------------------------------------------------------

export class SheetEvaluator {
  constructor(sheetModel, overrides = {}) {
    this.sheetModel = sheetModel; // { cells: {A1:{v,f}, ...} }
    this.overrides = Object.fromEntries(
      Object.entries(overrides).map(([k, v]) => [normalizeRef(k), v])
    );
    this.cache = new Map();
    this.astCache = new Map();
  }

  getCellRaw(ref) {
    const nref = normalizeRef(ref);
    if (Object.prototype.hasOwnProperty.call(this.overrides, nref)) return this.overrides[nref];
    const c = this.sheetModel.cells[nref] || this.sheetModel.cells[ref] || null;
    return c ? c.v : null;
  }

  getCellFormula(ref) {
    const nref = normalizeRef(ref);
    const c = this.sheetModel.cells[nref] || this.sheetModel.cells[ref] || null;
    return c ? c.f : null;
  }

  evalCell(ref, stack = []) {
    const nref = normalizeRef(ref);
    if (this.cache.has(nref)) return this.cache.get(nref);
    if (stack.includes(nref)) {
      // Avoid infinite loops; return blank.
      return "";
    }
    const f = this.getCellFormula(nref);
    if (!f) {
      const v = this.getCellRaw(nref);
      const out = isNumberLike(v) ? Number(v) : (v ?? "");
      this.cache.set(nref, out);
      return out;
    }

    let ast = this.astCache.get(f);
    if (!ast) {
      ast = parse(f);
      this.astCache.set(f, ast);
    }
    const out = this.evalAst(ast, stack.concat([nref]));
    this.cache.set(nref, out);
    return out;
  }

  evalAst(node, stack) {
    switch (node.type) {
      case "num":
        return node.value;
      case "str":
        return node.value;
      case "cell":
        return this.evalCell(node.ref, stack);
      case "range": {
        const refs = expandRange(node.ref);
        return refs.map((r) => this.evalCell(r, stack));
      }
      case "name":
        // In this template, treat unknown names as blank string.
        return "";
      case "unary": {
        const v = this.evalAst(node.expr, stack);
        if (node.op === "+") return toNumber(v);
        if (node.op === "-") return -toNumber(v);
        return 0;
      }
      case "bin": {
        const l = this.evalAst(node.left, stack);
        const r = this.evalAst(node.right, stack);
        switch (node.op) {
          case "&":
            return toText(l) + toText(r);
          case "+":
            return toNumber(l) + toNumber(r);
          case "-":
            return toNumber(l) - toNumber(r);
          case "*":
            return toNumber(l) * toNumber(r);
          case "/":
            return toNumber(l) / toNumber(r);
          case "=":
            return toText(l) === toText(r);
          case "<>":
            return toText(l) !== toText(r);
          case "<":
            return toNumber(l) < toNumber(r);
          case ">":
            return toNumber(l) > toNumber(r);
          case "<=":
            return toNumber(l) <= toNumber(r);
          case ">=":
            return toNumber(l) >= toNumber(r);
          default:
            return "";
        }
      }
      case "call": {
        const name = node.name;
        if (name === "IF") {
          const cond = this.evalAst(node.args[0], stack);
          return truthyExcel(cond)
            ? this.evalAst(node.args[1], stack)
            : this.evalAst(node.args[2] ?? { type: "str", value: "" }, stack);
        }
        if (name === "SUM") {
          let sum = 0;
          for (const arg of node.args) {
            const v = this.evalAst(arg, stack);
            if (Array.isArray(v)) {
              for (const x of v) sum += toNumber(x);
            } else {
              sum += toNumber(v);
            }
          }
          return sum;
        }
        if (name === "TRIM") {
          return trimExcel(this.evalAst(node.args[0], stack));
        }
        // Unknown function: blank
        return "";
      }
      default:
        return "";
    }
  }
}


