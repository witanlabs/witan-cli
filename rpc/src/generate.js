import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));

const schemaPath = join(__dirname, '..', 'schemas', 'witan-rpc.json');
const schema = JSON.parse(readFileSync(schemaPath, 'utf-8'));
const $defs = schema.$defs;

// ============================================================================
// Shared type renderer
// ============================================================================

function renderType(schema, lang, nameHint = '') {
  if (!schema || typeof schema !== 'object') {
    return lang === 'ts' ? 'unknown' : 'Any';
  }

  if (schema.$ref) {
    return schema.$ref.replace('#/$defs/', '');
  }

  if (schema.const !== undefined) {
    if (lang === 'ts') return JSON.stringify(schema.const);
    return `Literal[${JSON.stringify(schema.const)}]`;
  }

  if (Array.isArray(schema.enum)) {
    if (lang === 'ts') {
      return schema.enum.map((v) => JSON.stringify(v)).join(' | ');
    }
    return `Literal[${schema.enum.map((v) => JSON.stringify(v)).join(', ')}]`;
  }

  if (schema.type === 'null') {
    return lang === 'ts' ? 'null' : 'None';
  }

  if (schema.type === 'string') return lang === 'ts' ? 'string' : 'str';
  if (schema.type === 'integer') return lang === 'ts' ? 'number' : 'int';
  if (schema.type === 'number') return lang === 'ts' ? 'number' : 'float';
  if (schema.type === 'boolean') return lang === 'ts' ? 'boolean' : 'bool';

  if (schema.type === 'array') {
    const item = renderType(schema.items, lang, nameHint + 'Item');
    const needsParens = lang === 'ts' && item.includes('|');
    return lang === 'ts' ? `${needsParens ? `(${item})` : item}[]` : `list[${item}]`;
  }

  if (schema.type === 'object') {
    if (schema.additionalProperties && schema.additionalProperties !== true && !schema.properties) {
      const val = renderType(schema.additionalProperties, lang, nameHint + 'Value');
      return lang === 'ts' ? `Record<string, ${val}>` : `dict[str, ${val}]`;
    }
    if (schema.properties) {
      // Inline anonymous object - generate a name
      return `_inline_${nameHint || 'Object'}`;
    }
    return lang === 'ts' ? 'Record<string, unknown>' : 'dict[str, Any]';
  }

  if (Array.isArray(schema.anyOf)) {
    const parts = schema.anyOf.map((s, i) => renderType(s, lang, nameHint + 'Opt' + i));
    // Filter out duplicates and null/None pairs for cleaner output
    const seen = new Set();
    const unique = [];
    for (const p of parts) {
      const key = p;
      if (!seen.has(key)) {
        seen.add(key);
        unique.push(p);
      }
    }
    if (unique.length === 1) return unique[0];
    return unique.join(lang === 'ts' ? ' | ' : ' | ');
  }

  if (Array.isArray(schema.allOf)) {
    const refs = schema.allOf.filter((s) => s.$ref).map((s) => s.$ref.replace('#/$defs/', ''));
    const objects = schema.allOf.filter((s) => s.type === 'object' && s.properties);
    if (refs.length === 1 && objects.length === 1) {
      return refs[0]; // inheritance handled separately
    }
    return schema.allOf.map((s, i) => renderType(s, lang, nameHint + 'All' + i)).join(lang === 'ts' ? ' & ' : ' | ');
  }

  return lang === 'ts' ? 'unknown' : 'Any';
}

function collectInlineObjects(schema, nameHint, collected, lang) {
  if (!schema || typeof schema !== 'object') return;

  if (schema.type === 'object' && schema.properties) {
    const inlineName = `_inline_${nameHint}`;
    if (!collected.find((c) => c.name === inlineName)) {
      const props = {};
      const required = schema.required || [];
      for (const [k, v] of Object.entries(schema.properties)) {
        props[k] = v;
        collectInlineObjects(v, nameHint + k[0].toUpperCase() + k.slice(1), collected, lang);
      }
      collected.push({ name: inlineName, properties: props, required });
    }
    return;
  }

  if (schema.type === 'array') {
    collectInlineObjects(schema.items, nameHint + 'Item', collected, lang);
    return;
  }

  if (Array.isArray(schema.anyOf)) {
    schema.anyOf.forEach((s, i) => collectInlineObjects(s, nameHint + 'Opt' + i, collected, lang));
    return;
  }

  if (Array.isArray(schema.allOf)) {
    schema.allOf.forEach((s, i) => collectInlineObjects(s, nameHint + 'All' + i, collected, lang));
    return;
  }

  if (schema.additionalProperties && typeof schema.additionalProperties === 'object') {
    collectInlineObjects(schema.additionalProperties, nameHint + 'Value', collected, lang);
  }
}

// ============================================================================
// TypeScript Generator
// ============================================================================

function generateTypeScript() {
  const outFile = join(__dirname, '..', '..', 'node', 'src', 'generated-types.ts');

  const lines = [
    '// Auto-generated from rpc/schemas/witan-rpc.json',
    '// Do not edit manually. Run `npm run generate` in rpc/ to regenerate.',
    '',
  ];

  const aliasTypes = {
    ScalarCellValue: 'string | number | boolean | null',
  };

  for (const [name, tsType] of Object.entries(aliasTypes)) {
    lines.push(`export type ${name} = ${tsType};`);
  }
  lines.push('');

  const processed = new Set();

  function writeDef(name, def) {
    if (processed.has(name)) return;
    processed.add(name);

    if (aliasTypes[name]) return;

    // Simple type alias (enums, consts, primitive unions)
    const isAlias = !def.allOf && !def.properties &&
      (def.anyOf || def.enum || def.const);

    if (isAlias) {
      const ts = renderType(def, 'ts', name);
      lines.push(`export type ${name} = ${ts};`);
      lines.push('');
      return;
    }

    // $ref alias (e.g. ChartLineFormatSpec = ChartBorderFormatSpec)
    if (def.$ref) {
      lines.push(`export type ${name} = ${def.$ref.replace('#/$defs/', '')};`);
      lines.push('');
      return;
    }

    let parentNames = [];
    let ownProps = def.properties || {};
    let ownRequired = def.required || [];

    if (def.allOf) {
      for (const part of def.allOf) {
        if (part.$ref) {
          parentNames.push(part.$ref.replace('#/$defs/', ''));
        } else if (part.type === 'object') {
          ownProps = { ...ownProps, ...part.properties };
          ownRequired = [...new Set([...ownRequired, ...(part.required || [])])];
        }
      }
    }

    // Collect inline objects
    const inlineObjects = [];
    for (const [propName, propSchema] of Object.entries(ownProps)) {
      collectInlineObjects(propSchema, name + propName[0].toUpperCase() + propName.slice(1), inlineObjects, 'ts');
    }
    for (const inline of inlineObjects) {
      if (processed.has(inline.name)) continue;
      processed.add(inline.name);
      lines.push(`interface ${inline.name} {`);
      for (const [pn, ps] of Object.entries(inline.properties)) {
        const opt = !inline.required.includes(pn);
        const ptype = renderType(ps, 'ts', inline.name + pn[0].toUpperCase() + pn.slice(1));
        lines.push(`  ${pn}${opt ? '?' : ''}: ${ptype};`);
      }
      lines.push('}');
      lines.push('');
    }

    const extendsClause = parentNames.length > 0 ? ` extends ${parentNames.join(', ')}` : '';
    lines.push(`export interface ${name}${extendsClause} {`);

    const propNames = Object.keys(ownProps);
    if (propNames.length === 0) {
      lines.push('  // no own properties');
    } else {
      for (const [propName, propSchema] of Object.entries(ownProps)) {
        const isOptional = !ownRequired.includes(propName);
        const tsType = renderType(propSchema, 'ts', name + propName[0].toUpperCase() + propName.slice(1));
        lines.push(`  ${propName}${isOptional ? '?' : ''}: ${tsType};`);
      }
    }
    lines.push('}');
    lines.push('');
  }

  // Topological sort
  const deps = new Map();
  function collectDeps(def) {
    const result = new Set();
    if (!def || typeof def !== 'object') return result;
    if (def.$ref) {
      result.add(def.$ref.replace('#/$defs/', ''));
      return result;
    }
    if (def.type === 'array' && def.items) {
      for (const d of collectDeps(def.items)) result.add(d);
    }
    if (def.type === 'object' && def.additionalProperties && typeof def.additionalProperties === 'object') {
      for (const d of collectDeps(def.additionalProperties)) result.add(d);
    }
    if (def.properties) {
      for (const p of Object.values(def.properties)) {
        for (const d of collectDeps(p)) result.add(d);
      }
    }
    if (def.anyOf) {
      for (const p of def.anyOf) {
        for (const d of collectDeps(p)) result.add(d);
      }
    }
    if (def.allOf) {
      for (const p of def.allOf) {
        for (const d of collectDeps(p)) result.add(d);
      }
    }
    return result;
  }

  for (const [name, def] of Object.entries($defs)) {
    deps.set(name, collectDeps(def));
  }

  const order = [];
  function visit(name) {
    if (processed.has(name)) return;
    if (order.includes(name)) return;
    for (const dep of deps.get(name) || []) {
      if (dep !== name) visit(dep);
    }
    order.push(name);
  }
  for (const name of Object.keys($defs)) {
    visit(name);
  }

  for (const name of order) {
    writeDef(name, $defs[name]);
  }

  writeFileSync(outFile, lines.join('\n'));
  console.log(`Generated TypeScript types in ${outFile}`);
}

// ============================================================================
// Python Generator
// ============================================================================

function generatePython() {
  const outFile = join(__dirname, '..', '..', 'python', 'witan', 'generated_types.py');

  const lines = [
    '# Auto-generated from rpc/schemas/witan-rpc.json',
    '# Do not edit manually. Run `npm run generate` in rpc/ to regenerate.',
    '',
    'from __future__ import annotations',
    '',
    'from typing import Any, Literal, TypeAlias, TypedDict',
    'from typing_extensions import NotRequired',
    '',
    'ScalarCellValue: TypeAlias = str | int | float | bool | None',
    '',
  ];

  const processed = new Set();

  function writeDef(name, def) {
    if (processed.has(name)) return;
    processed.add(name);

    if (name === 'ScalarCellValue') {
      return;
    }

    // $ref alias
    if (def.$ref) {
      const target = def.$ref.replace('#/$defs/', '');
      lines.push(`${name}: TypeAlias = ${target}`);
      lines.push('');
      return;
    }

    const isAlias = !def.allOf && !def.properties &&
      (def.anyOf || def.enum || def.const);

    if (isAlias) {
      const py = renderType(def, 'py', name);
      lines.push(`${name}: TypeAlias = ${py}`);
      lines.push('');
      return;
    }

    let parentNames = [];
    let ownProps = def.properties || {};
    let ownRequired = def.required || [];

    if (def.allOf) {
      for (const part of def.allOf) {
        if (part.$ref) {
          parentNames.push(part.$ref.replace('#/$defs/', ''));
        } else if (part.type === 'object') {
          ownProps = { ...ownProps, ...part.properties };
          ownRequired = [...new Set([...ownRequired, ...(part.required || [])])];
        }
      }
    }

    // Collect inline objects
    const inlineObjects = [];
    for (const [propName, propSchema] of Object.entries(ownProps)) {
      collectInlineObjects(propSchema, name + propName[0].toUpperCase() + propName.slice(1), inlineObjects, 'py');
    }
    for (const inline of inlineObjects) {
      if (processed.has(inline.name)) continue;
      processed.add(inline.name);
      const inlineProps = Object.keys(inline.properties);
      const inlineAllOpt = inlineProps.length > 0 && inlineProps.every((p) => !inline.required.includes(p));
      const inlineHasKeyword = inlineProps.some((p) => pythonKeywords.has(p));
      if (inlineHasKeyword) {
        const totalKw = inlineAllOpt ? ', total=False' : '';
        lines.push(`${inline.name} = TypedDict("${inline.name}", {`);
        for (const [pn, ps] of Object.entries(inline.properties)) {
          const opt = !inline.required.includes(pn);
          const ptype = renderType(ps, 'py', inline.name + pn[0].toUpperCase() + pn.slice(1));
          const annotation = opt ? `NotRequired[${ptype}]` : ptype;
          lines.push(`    "${pn}": ${annotation},`);
        }
        lines.push(`}${totalKw})`);
        lines.push('');
      } else {
        if (inlineAllOpt) {
          lines.push(`class ${inline.name}(TypedDict, total=False):`);
        } else {
          lines.push(`class ${inline.name}(TypedDict):`);
        }
        for (const [pn, ps] of Object.entries(inline.properties)) {
          const opt = !inline.required.includes(pn);
          const ptype = renderType(ps, 'py', inline.name + pn[0].toUpperCase() + pn.slice(1));
          const annotation = opt ? `NotRequired[${ptype}]` : ptype;
          lines.push(`    ${pn}: ${annotation}`);
        }
        lines.push('');
      }
    }

    const propNames = Object.keys(ownProps);
    const allOptional = propNames.length > 0 && propNames.every((p) => !ownRequired.includes(p));

    const pythonKeywords = new Set([
      'False','None','True','and','as','assert','async','await','break','class','continue',
      'def','del','elif','else','except','finally','for','from','global','if','import',
      'in','is','lambda','nonlocal','not','or','pass','raise','return','try','while',
      'with','yield',
    ]);
    const hasKeywordField = propNames.some((p) => pythonKeywords.has(p));

    if (hasKeywordField) {
      // Use TypedDict("Name", {...}) form to allow keyword field names
      const totalKw = allOptional && parentNames.length === 0 ? ', total=False' : '';
      const entries = [];
      for (const [propName, propSchema] of Object.entries(ownProps)) {
        const isOptional = !ownRequired.includes(propName);
        const pyAnn = renderType(propSchema, 'py', name + propName[0].toUpperCase() + propName.slice(1));
        const annotation = isOptional ? `NotRequired[${pyAnn}]` : pyAnn;
        entries.push(`    "${propName}": ${annotation}`);
      }
      if (parentNames.length > 0) {
        lines.push(`${name} = TypedDict("${name}", {`);
      } else {
        lines.push(`${name} = TypedDict("${name}", {`);
      }
      if (def.description) {
        lines.push(`    # ${def.description}`);
      }
      for (const entry of entries) {
        lines.push(entry + ',');
      }
      lines.push(`}${totalKw})`);
      lines.push('');
    } else {
      let baseClause = '';
      if (parentNames.length > 0) {
        baseClause = `(${parentNames.join(', ')})`;
      } else if (allOptional) {
        baseClause = '(TypedDict, total=False)';
      } else {
        baseClause = '(TypedDict)';
      }

      lines.push(`class ${name}${baseClause}:`);

      if (def.description) {
        lines.push(`    """${def.description}"""`);
      }

      if (propNames.length === 0) {
        lines.push('    pass');
      } else {
        for (const [propName, propSchema] of Object.entries(ownProps)) {
          const isOptional = !ownRequired.includes(propName);
          const pyAnn = renderType(propSchema, 'py', name + propName[0].toUpperCase() + propName.slice(1));
          const annotation = isOptional ? `NotRequired[${pyAnn}]` : pyAnn;
          lines.push(`    ${propName}: ${annotation}`);
        }
      }
      lines.push('');
    }
  }

  // Topological sort
  const deps = new Map();
  function collectDeps(def) {
    const result = new Set();
    if (!def || typeof def !== 'object') return result;
    if (def.$ref) {
      result.add(def.$ref.replace('#/$defs/', ''));
      return result;
    }
    if (def.type === 'array' && def.items) {
      for (const d of collectDeps(def.items)) result.add(d);
    }
    if (def.type === 'object' && def.additionalProperties && typeof def.additionalProperties === 'object') {
      for (const d of collectDeps(def.additionalProperties)) result.add(d);
    }
    if (def.properties) {
      for (const p of Object.values(def.properties)) {
        for (const d of collectDeps(p)) result.add(d);
      }
    }
    if (def.anyOf) {
      for (const p of def.anyOf) {
        for (const d of collectDeps(p)) result.add(d);
      }
    }
    if (def.allOf) {
      for (const p of def.allOf) {
        for (const d of collectDeps(p)) result.add(d);
      }
    }
    return result;
  }

  for (const [name, def] of Object.entries($defs)) {
    deps.set(name, collectDeps(def));
  }

  const order = [];
  function visit(name) {
    if (processed.has(name)) return;
    if (order.includes(name)) return;
    for (const dep of deps.get(name) || []) {
      if (dep !== name) visit(dep);
    }
    order.push(name);
  }
  for (const name of Object.keys($defs)) {
    visit(name);
  }

  for (const name of order) {
    writeDef(name, $defs[name]);
  }

  writeFileSync(outFile, lines.join('\n'));
  console.log(`Generated Python types in ${outFile}`);
}

// ============================================================================
// Run
// ============================================================================

function main() {
  generateTypeScript();
  generatePython();
}

main();
