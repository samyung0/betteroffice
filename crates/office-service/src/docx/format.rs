//! yrs/saveFormatting.ts: the pPr, tblPr, trPr and tcPr a projection saves.

use crate::js::utf16_compare;
use crate::jsv::{Obj, V, quoted, stringify};
use crate::obj;

use super::item::o;

/// saveFormatting.ts `canonical`: sorted keys, nulls dropped.
fn canonical(value: &V) -> String {
    match value {
        V::Arr(items) => format!(
            "[{}]",
            items
                .items()
                .iter()
                .map(canonical)
                .collect::<Vec<_>>()
                .join(",")
        ),
        V::Obj(object) => {
            let mut entries: Vec<_> = object
                .entries()
                .into_iter()
                .filter(|(_, entry)| !entry.nullish())
                .collect();
            entries.sort_by(|(a, _), (b, _)| utf16_compare(a, b));
            format!(
                "{{{}}}",
                entries
                    .iter()
                    .map(|(key, entry)| format!("{}:{}", quoted(key), canonical(entry)))
                    .collect::<Vec<_>>()
                    .join(",")
            )
        }
        V::Undef => "null".into(),
        other => stringify(other).unwrap_or_else(|| "null".into()),
    }
}

/// The pPr properties the editor holds resolved (direct, else list level, else style).
const RESOLVED_PROPERTIES: [&str; 26] = [
    "alignment",
    "spaceBefore",
    "spaceAfter",
    "spaceBeforeLines",
    "spaceAfterLines",
    "beforeAutospacing",
    "afterAutospacing",
    "lineSpacing",
    "lineSpacingRule",
    "indentLeft",
    "indentRight",
    "indentFirstLine",
    "hangingIndent",
    "borders",
    "shading",
    "tabs",
    "pageBreakBefore",
    "keepNext",
    "keepLines",
    "widowControl",
    "contextualSpacing",
    "snapToGrid",
    "autoSpaceDE",
    "autoSpaceDN",
    "outlineLevel",
    "bidi",
];

/// Properties one OOXML attribute set carries together.
const LINKED: [[&str; 2]; 2] = [
    ["lineSpacing", "lineSpacingRule"],
    ["indentFirstLine", "hangingIndent"],
];

/// Character-unit twins Word reads in place of a twips indent.
fn character_twins(key: &str) -> &'static [&'static str] {
    match key {
        "indentLeft" => &["indentLeftChars"],
        "indentRight" => &["indentRightChars"],
        "indentFirstLine" => &["indentFirstLineChars", "hangingIndentChars"],
        _ => &[],
    }
}

/// Tab stops in the model's shape; editor ops store `{pos, val}`.
fn model_tabs(value: &V) -> V {
    let Some(stops) = value.arr() else {
        return V::Undef;
    };
    V::array(
        stops
            .items()
            .iter()
            .map(|stop| {
                let tab = obj! {
                    "position": stop.get("position").or_else(|| stop.get("pos")),
                    "alignment": stop.get("alignment").or_else(|| stop.get("val")),
                };
                if !stop.get("leader").nullish()
                    && let Some(object) = tab.obj()
                {
                    object.set("leader", stop.get("leader"));
                }
                tab
            })
            .collect(),
    )
}

fn model_value(attrs: &V, key: &str) -> V {
    if key == "tabs" {
        model_tabs(&attrs.get("tabs"))
    } else {
        attrs.get(key)
    }
}

/// The pPr a paragraph saves with: its source formatting with every resolved
/// property the editor holds differently from its seed written over it.
pub(crate) fn paragraph_attrs_to_formatting(attrs: &V, seeded: &V) -> V {
    let orig = attrs
        .get("_originalFormatting")
        .or_else(|| V::Obj(Obj::new()));
    let result = o(&orig).map(|orig| orig.spread()).unwrap_or_default();
    let mut changed: Vec<&str> = RESOLVED_PROPERTIES
        .into_iter()
        .filter(|key| {
            let seed = if *key == "tabs" {
                model_tabs(&seeded.get("tabs"))
            } else {
                seeded.get(key)
            };
            canonical(&model_value(attrs, key)) != canonical(&seed)
        })
        .collect();
    for group in LINKED {
        if group.iter().any(|key| changed.contains(key)) {
            for key in group {
                if !changed.contains(&key) {
                    changed.push(key);
                }
            }
        }
    }
    for key in changed {
        let value = model_value(attrs, key);
        if value.nullish() {
            result.delete(key);
        } else {
            result.set(key, value);
        }
        for twin in character_twins(key) {
            result.delete(twin);
        }
    }
    let num_pr = attrs.get("numPr");
    let from_style = attrs.get("numPrFromStyle");
    if !from_style.nullish() && !num_pr.nullish() && stringify(&num_pr) == stringify(&from_style) {
        result.delete("numPr");
        result.delete("numPrFromStyle");
    } else if !num_pr.same(&orig.get("numPr"))
        && stringify(&num_pr) != stringify(&orig.get("numPr"))
    {
        result.set("numPr", num_pr.or(|| V::Undef));
        result.delete("numPrFromStyle");
    }
    let style = attrs.get("styleId");
    if !style.same(&orig.get("styleId").or(|| V::Undef)) {
        result.set("styleId", style.or(|| V::Undef));
    }
    if result.values().iter().any(|value| !value.is_undef()) {
        V::Obj(result)
    } else {
        V::Undef
    }
}

fn width_value(attrs: &V) -> V {
    let width = attrs.get("width");
    if !width.nullish() || attrs.get("widthType").truthy() {
        obj! {
            "value": width.or_else(|| V::Num(0.0)),
            "type": attrs.get("widthType").or(|| V::str("dxa")),
        }
    } else {
        V::Undef
    }
}

fn measurement_margins(margins: &V) -> V {
    let side = |key: &str| {
        let value = margins.get(key);
        if value.nullish() {
            V::Undef
        } else {
            obj! { "value": value, "type": "dxa" }
        }
    };
    obj! { "top": side("top"), "bottom": side("bottom"), "left": side("left"), "right": side("right") }
}

/// `attrs[key] !== (orig[key] || undefined)`, then `result[key] = attrs[key] || undefined`.
fn changed_to(result: &Obj, attrs: &V, orig: &V, attr: &str, field: &str) {
    if !attrs.get(attr).same(&orig.get(field).or(|| V::Undef)) {
        result.set(field, attrs.get(attr).or(|| V::Undef));
    }
}

pub(crate) fn table_attrs_to_formatting(attrs: &V) -> V {
    let orig = attrs.get("_originalFormatting");
    if orig.truthy() {
        let result = o(&orig).map(|orig| orig.spread()).unwrap_or_default();
        changed_to(&result, attrs, &orig, "styleId", "styleId");
        changed_to(&result, attrs, &orig, "justification", "justification");
        changed_to(&result, attrs, &orig, "floating", "floating");
        changed_to(&result, attrs, &orig, "tableLayout", "layout");
        changed_to(&result, attrs, &orig, "look", "look");
        changed_to(&result, attrs, &orig, "bidi", "bidi");
        let width = orig.get("width");
        if !attrs.get("width").same(&width.get("value"))
            || !attrs.get("widthType").same(&width.get("type"))
        {
            result.set("width", width_value(attrs));
        }
        if attrs.get("cellMargins").truthy() {
            result.set(
                "cellMargins",
                measurement_margins(&attrs.get("cellMargins")),
            );
        }
        return V::Obj(result);
    }
    let has_formatting = attrs.get("styleId").truthy()
        || !attrs.get("width").nullish()
        || [
            "widthType",
            "justification",
            "tableLayout",
            "floating",
            "cellMargins",
            "look",
            "bidi",
        ]
        .iter()
        .any(|key| attrs.get(key).truthy());
    if !has_formatting {
        return V::Undef;
    }
    let or_undef = |key: &str| attrs.get(key).or(|| V::Undef);
    obj! {
        "styleId": or_undef("styleId"),
        "width": width_value(attrs),
        "justification": or_undef("justification"),
        "layout": or_undef("tableLayout"),
        "floating": or_undef("floating"),
        "cellMargins": if attrs.get("cellMargins").truthy() {
            measurement_margins(&attrs.get("cellMargins"))
        } else {
            V::Undef
        },
        "look": or_undef("look"),
        "bidi": or_undef("bidi"),
    }
}

pub(crate) fn table_row_attrs_to_formatting(attrs: &V) -> V {
    let height = |attrs: &V| {
        let height = attrs.get("height");
        if height.truthy() {
            obj! { "value": height, "type": "dxa" }
        } else {
            V::Undef
        }
    };
    let orig = attrs.get("_originalFormatting");
    if orig.truthy() {
        let result = o(&orig).map(|orig| orig.spread()).unwrap_or_default();
        if !attrs
            .get("height")
            .same(&orig.get("height").get("value").or(|| V::Undef))
        {
            result.set("height", height(attrs));
        }
        changed_to(&result, attrs, &orig, "heightRule", "heightRule");
        changed_to(&result, attrs, &orig, "isHeader", "header");
        return V::Obj(result);
    }
    if !attrs.get("height").truthy() && !attrs.get("isHeader").truthy() {
        return V::Undef;
    }
    obj! {
        "height": height(attrs),
        "heightRule": attrs.get("heightRule").or(|| V::Undef),
        "header": attrs.get("isHeader").or(|| V::Undef),
    }
}

fn cell_margins(margins: &V) -> V {
    let result = Obj::new();
    for key in ["top", "bottom", "left", "right"] {
        let value = margins.get(key);
        if !value.nullish() {
            result.set(key, obj! { "value": value, "type": "dxa" });
        }
    }
    V::Obj(result)
}

/// The sides of a cell's resolved borders or margins it saves: for a source
/// cell only those that differ from what the seed resolved (`seeded`), the
/// cell's own tcPr keeping the rest, so a table style or table-wide margins
/// keep applying in Word (Epo 2026-10-21); all of them for any other cell.
fn changed_sides(live: &V, seeded: Option<&V>) -> V {
    let Some(seeded) = seeded else {
        return live.clone();
    };
    let changed = Obj::new();
    for (side, value) in o(live).map(|live| live.entries()).unwrap_or_default() {
        if canonical(&value) != canonical(&seeded.get(&side)) {
            changed.set(&side, value);
        }
    }
    if changed.is_empty() {
        V::Undef
    } else {
        V::Obj(changed)
    }
}

/// `sides` written over the source cell's own `borders` or `margins`.
fn over_source(source: &V, sides: V) -> V {
    let merged = o(source).map(|source| source.spread()).unwrap_or_default();
    if let Some(sides) = o(&sides) {
        merged.assign(&sides);
    }
    V::Obj(merged)
}

pub(crate) fn table_cell_attrs_to_formatting(attrs: &V, seeded: Option<&(V, V)>) -> V {
    let cell_width = |attrs: &V| {
        obj! {
            "value": attrs.get("width"),
            "type": attrs.get("widthType").or(|| V::str("dxa")),
        }
    };
    let colspan = attrs.get("colspan").to_number();
    let borders = changed_sides(&attrs.get("borders"), seeded.map(|(borders, _)| borders));
    let margins = changed_sides(&attrs.get("margins"), seeded.map(|(_, margins)| margins));
    let orig = attrs.get("_originalFormatting");
    if orig.truthy() {
        let result = o(&orig).map(|orig| orig.spread()).unwrap_or_default();
        if colspan > 1.0 {
            result.set("gridSpan", attrs.get("colspan"));
        }
        if !attrs.get("width").nullish() {
            result.set("width", cell_width(attrs));
        }
        changed_to(&result, attrs, &orig, "verticalAlign", "verticalAlign");
        let background = attrs.get("backgroundColor");
        if background.truthy() {
            let shading = orig.get("shading");
            result.set(
                "shading",
                if attrs.get("_originalResolvedFill").same(&background) && shading.truthy() {
                    shading
                } else {
                    obj! { "fill": obj! { "rgb": background } }
                },
            );
        } else if orig.get("shading").truthy() {
            result.set("shading", V::Undef);
        }
        if borders.truthy() {
            result.set(
                "borders",
                if seeded.is_some() {
                    over_source(&orig.get("borders"), borders)
                } else {
                    borders
                },
            );
        }
        if margins.truthy() {
            let sides = cell_margins(&margins);
            result.set(
                "margins",
                if seeded.is_some() {
                    over_source(&orig.get("margins"), sides)
                } else {
                    sides
                },
            );
        }
        changed_to(&result, attrs, &orig, "textDirection", "textDirection");
        if attrs.get("noWrap").truthy() != orig.get("noWrap").truthy() {
            result.set("noWrap", attrs.get("noWrap").or(|| V::Undef));
        }
        return V::Obj(result);
    }
    let has_formatting = colspan > 1.0
        || attrs.get("rowspan").to_number() > 1.0
        || !attrs.get("width").nullish()
        || borders.truthy()
        || margins.truthy()
        || [
            "verticalAlign",
            "backgroundColor",
            "textDirection",
            "noWrap",
        ]
        .iter()
        .any(|key| attrs.get(key).truthy());
    if !has_formatting {
        return V::Undef;
    }
    obj! {
        "gridSpan": if colspan > 1.0 { attrs.get("colspan") } else { V::Undef },
        "width": if attrs.get("width").nullish() { V::Undef } else { cell_width(attrs) },
        "verticalAlign": attrs.get("verticalAlign").or(|| V::Undef),
        "textDirection": attrs.get("textDirection").or(|| V::Undef),
        "shading": if attrs.get("backgroundColor").truthy() {
            obj! { "fill": obj! { "rgb": attrs.get("backgroundColor") } }
        } else {
            V::Undef
        },
        "borders": borders,
        "margins": if margins.truthy() { cell_margins(&margins) } else { V::Undef },
        "noWrap": attrs.get("noWrap").or(|| V::Undef),
    }
}
