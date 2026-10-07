//! yrsToDocument.ts, part one: a story's inline items (text runs and
//! embeds, as `{kind, text|embedKind, payload, attributes}` objects) to the
//! serializer's paragraph content. Ported line by line on JS values, since
//! the TS mutates and compares those objects by reference.

use std::collections::HashMap;

use crate::error::{Error, Result};
use crate::js::utf16_compare;
use crate::jsv::{Arr, Obj, Str, V, quoted, stringify, text};
use crate::{arr, obj};

use super::sdt::{apply_content_control_value, sdt_attrs_to_props};

pub(crate) const BOOLEAN_MARKS: [&str; 12] = [
    "bold",
    "italic",
    "superscript",
    "subscript",
    "allCaps",
    "smallCaps",
    "emboss",
    "imprint",
    "textShadow",
    "textOutline",
    "hidden",
    "rtl",
];

pub(crate) fn o(value: &V) -> Option<Obj> {
    value.obj().cloned()
}

pub(crate) fn st(value: &V) -> Option<Str> {
    value.as_str()
}

/// `value.type`.
pub(crate) fn ty(value: &V) -> Str {
    value.get("type").as_str().unwrap_or_else(|| Str::from(""))
}

pub(crate) fn is(value: &V, kind: &str) -> bool {
    &*ty(value) == kind
}

/// `Math.round`.
pub(crate) fn js_round(value: f64) -> f64 {
    let floor = value.floor();
    if value - floor >= 0.5 {
        floor + 1.0
    } else {
        floor
    }
}

pub(crate) fn px_to_emu(px: f64) -> f64 {
    js_round(px / 96.0 * 914_400.0)
}

/// `asString(value) || fallback`.
fn str_or(value: &V, fallback: &str) -> V {
    match st(value) {
        Some(text) if !text.is_empty() => V::Str(text),
        _ => V::str(fallback),
    }
}

/// `asString(value) || undefined`.
fn str_or_undef(value: &V) -> V {
    match st(value) {
        Some(text) if !text.is_empty() => V::Str(text),
        _ => V::Undef,
    }
}

pub(crate) fn drop_nulls(value: &V) -> V {
    if let Some(items) = value.arr() {
        return V::array(items.items().iter().map(drop_nulls).collect());
    }
    let Some(object) = value.obj() else {
        return value.clone();
    };
    let result = Obj::new();
    for (key, entry) in object.entries() {
        if entry.nullish() {
            continue;
        }
        result.set(&key, drop_nulls(&entry));
    }
    V::Obj(result)
}

/// yrsToDocument.ts `stableStringify`: sorted keys, `undefined` as null.
pub(crate) fn stable_stringify(value: &V) -> String {
    match value {
        V::Undef | V::Null => "null".into(),
        V::Arr(items) => format!(
            "[{}]",
            items
                .items()
                .iter()
                .map(stable_stringify)
                .collect::<Vec<_>>()
                .join(",")
        ),
        V::Obj(object) => {
            let mut keys = object.keys();
            keys.sort_by(|a, b| utf16_compare(a, b));
            format!(
                "{{{}}}",
                keys.iter()
                    .map(|key| format!("{}:{}", quoted(key), stable_stringify(&object.get(key))))
                    .collect::<Vec<_>>()
                    .join(",")
            )
        }
        other => stringify(other).unwrap_or_default(),
    }
}

pub(crate) fn formatting_attrs(attributes: &V) -> V {
    let result = o(attributes)
        .map(|object| object.spread())
        .unwrap_or_default();
    result.delete("hyperlink");
    result.delete("ins");
    result.delete("del");
    V::Obj(result)
}

/// Converts attributes without manufacturing mark objects.
pub(crate) fn attrs_to_text_formatting(attributes: &V) -> Obj {
    let formatting = Obj::new();
    let get = |key: &str| attributes.get(key);
    if get("bold").truthy() {
        formatting.set("bold", V::Bool(true));
        formatting.set("boldCs", V::Bool(true));
    }
    if get("italic").truthy() {
        formatting.set("italic", V::Bool(true));
        formatting.set("italicCs", V::Bool(true));
    }
    if let Some(underline) = o(&get("underline"))
        && !underline.get("inheritedHyperlink").same(&V::Bool(true))
    {
        formatting.set(
            "underline",
            obj! {
                "style": str_or(&underline.get("style"), "single"),
                "color": o(&underline.get("color")).map(V::Obj).unwrap_or(V::Null),
            },
        );
    }
    if let Some(strike) = o(&get("strike")) {
        if strike.get("double").truthy() {
            formatting.set("doubleStrike", V::Bool(true));
        } else {
            formatting.set("strike", V::Bool(true));
        }
    }
    if let Some(color) = o(&get("textColor"))
        && !color.get("inheritedHyperlink").same(&V::Bool(true))
    {
        let or_null = |value: Option<Str>| value.map(V::Str).unwrap_or(V::Null);
        formatting.set(
            "color",
            obj! {
                "rgb": or_null(st(&color.get("rgb"))),
                "themeColor": color.get("themeColor").or_else(|| V::Null),
                "themeTint": or_null(st(&color.get("themeTint"))),
                "themeShade": or_null(st(&color.get("themeShade"))),
            },
        );
    }
    if let V::Str(highlight) = get("highlight") {
        formatting.set("highlight", V::Str(highlight));
    }
    if let Some(font_size) = o(&get("fontSize")) {
        let size = font_size.get("size").finite();
        let size_cs = font_size.get("sizeCs").finite().or(size);
        if let Some(size) = size {
            formatting.set("fontSize", V::Num(size));
        }
        if let Some(size_cs) = size_cs {
            formatting.set("fontSizeCs", V::Num(size_cs));
        }
    }
    if let Some(language) = o(&get("language")) {
        formatting.set(
            "language",
            obj! {
                "latin": st(&language.get("latin")),
                "eastAsia": st(&language.get("eastAsia")),
                "bidi": st(&language.get("bidi")),
            },
        );
    }
    if let Some(family) = o(&get("fontFamily")) {
        // A slot's theme font wins over its name in Word; the seed resolves the
        // name from the theme, which the source did not write.
        let font = |slot: &str| {
            if family.get(&format!("{slot}Theme")).truthy() {
                None
            } else {
                st(&family.get(slot))
            }
        };
        let or_null = |value: Option<Str>| value.map(V::Str).unwrap_or(V::Null);
        formatting.set(
            "fontFamily",
            obj! {
                "ascii": or_null(font("ascii")),
                "hAnsi": or_null(font("hAnsi")),
                "eastAsia": font("eastAsia"),
                "cs": font("cs"),
                "asciiTheme": family.get("asciiTheme").or_else(|| V::Null),
                "hAnsiTheme": st(&family.get("hAnsiTheme")),
                "eastAsiaTheme": st(&family.get("eastAsiaTheme")),
                "csTheme": st(&family.get("csTheme")),
            },
        );
    }
    if let Some(hint) = st(&get("fontHint").get("hint")).filter(|hint| !hint.is_empty()) {
        let family = o(&formatting.get("fontFamily"))
            .map(|family| family.spread())
            .unwrap_or_default();
        family.set("hint", V::Str(hint));
        formatting.set("fontFamily", V::Obj(family));
    }
    if get("superscript").truthy() {
        formatting.set("vertAlign", V::str("superscript"));
    }
    if get("subscript").truthy() {
        formatting.set("vertAlign", V::str("subscript"));
    }
    if get("allCaps").truthy() {
        formatting.set("allCaps", V::Bool(true));
    }
    if get("smallCaps").truthy() {
        formatting.set("smallCaps", V::Bool(true));
    }
    if let Some(spacing) = o(&get("characterSpacing")) {
        for (from, to) in [
            ("spacing", "spacing"),
            ("position", "position"),
            ("scale", "scale"),
            ("kerning", "kerning"),
        ] {
            if let Some(value) = spacing.get(from).finite() {
                formatting.set(to, V::Num(value));
            }
        }
    }
    if get("emboss").truthy() {
        formatting.set("emboss", V::Bool(true));
    }
    if get("imprint").truthy() {
        formatting.set("imprint", V::Bool(true));
    }
    if get("textShadow").truthy() {
        formatting.set("shadow", V::Bool(true));
    }
    if let Some(emphasis) = o(&get("emphasisMark")) {
        formatting.set("emphasisMark", str_or(&emphasis.get("type"), "dot"));
    }
    if get("textOutline").truthy() {
        formatting.set("outline", V::Bool(true));
    }
    if get("hidden").truthy() {
        formatting.set("hidden", V::Bool(true));
    }
    if get("rtl").truthy() {
        formatting.set("rtl", V::Bool(true));
    }
    if let Some(effect) = o(&get("textEffect")) {
        formatting.set("effect", str_or(&effect.get("effect"), "blinkBackground"));
    }
    let modern = get("modernTextEffects").get("effects");
    if modern.truthy() {
        formatting.set("modernEffects", modern);
    }
    if let Some(style) = st(&get("runStyle").get("styleId")).filter(|style| !style.is_empty()) {
        formatting.set("styleId", V::Str(style));
    }
    formatting
}

pub(crate) fn run_content_for_text(text: &str, formatting: &Obj) -> Vec<V> {
    let mut content = Vec::new();
    let mut plain = String::new();
    let flush = |content: &mut Vec<V>, plain: &mut String| {
        if !plain.is_empty() {
            content.push(obj! { "type": "text", "text": std::mem::take(plain) });
        }
    };
    let family = formatting.get("fontFamily");
    let fonts: Vec<Str> = if family.truthy() {
        ["ascii", "hAnsi", "eastAsia", "cs"]
            .iter()
            .filter_map(|slot| family.get(slot).as_str().filter(|font| !font.is_empty()))
            .collect()
    } else {
        Vec::new()
    };
    let symbol_font =
        (!fonts.is_empty() && fonts.iter().all(|font| font == &fonts[0])).then(|| fonts[0].clone());
    for ch in text.chars() {
        let code = ch as u32;
        if ch == '\u{ad}' {
            flush(&mut content, &mut plain);
            content.push(obj! { "type": "softHyphen" });
        } else if ch == '\u{2011}' {
            flush(&mut content, &mut plain);
            content.push(obj! { "type": "noBreakHyphen" });
        } else if let Some(font) = symbol_font
            .as_ref()
            .filter(|_| (0xf000..=0xf8ff).contains(&code))
        {
            flush(&mut content, &mut plain);
            content.push(obj! {
                "type": "symbol",
                "font": V::Str(font.clone()),
                "char": format!("{code:04X}"),
            });
        } else {
            plain.push(ch);
        }
    }
    flush(&mut content, &mut plain);
    content
}

pub(crate) fn run_property_changes(attributes: &V) -> V {
    let changes = attributes.get("rPrChange").get("changes");
    let Some(changes) = changes.arr() else {
        return V::Undef;
    };
    V::array(
        changes
            .items()
            .iter()
            .map(|change| {
                let copy = o(change).map(|change| change.spread()).unwrap_or_default();
                copy.delete("previousAttributes");
                copy.delete("currentAttributes");
                V::Obj(copy)
            })
            .collect(),
    )
}

pub(crate) fn create_text_run(text: &str, attributes: &V) -> V {
    let formatting = attrs_to_text_formatting(&formatting_attrs(attributes));
    let content = run_content_for_text(text, &formatting);
    obj! {
        "type": "run",
        "formatting": if formatting.is_empty() { V::Undef } else { V::Obj(formatting) },
        "propertyChanges": run_property_changes(attributes),
        "content": V::array(content),
    }
}

pub(crate) fn append_text_run(target: &V, source: &V) {
    let Some(content) = target.get("content").arr().cloned() else {
        return;
    };
    for item in source.get("content").items() {
        let previous = content.last();
        if is(&previous, "text") && is(&item, "text") {
            let joined = format!(
                "{}{}",
                previous.get("text").to_js_string(),
                item.get("text").to_js_string()
            );
            if let Some(previous) = previous.obj() {
                previous.set("text", V::from(joined));
            }
        } else {
            content.push(item);
        }
    }
}

fn fnv32(value: &str) -> u32 {
    let mut hash: u32 = 0x811c_9dc5;
    for byte in value.bytes() {
        hash = (hash ^ byte as u32).wrapping_mul(0x0100_0193);
    }
    hash
}

/// The `w:id` a revision id saves as (`revisionId`).
pub(crate) fn revision_id(value: &V) -> f64 {
    match value {
        V::Num(number) if number.is_finite() => *number,
        V::Str(text) => {
            let parsed = value.to_number();
            if parsed.is_finite() {
                parsed
            } else {
                (1u64 << 30) as f64 + (fnv32(text) % (1 << 30)) as f64
            }
        }
        _ => 0.0,
    }
}

pub(crate) fn tracked_info(raw: &V) -> Option<V> {
    let value = o(raw)?;
    let author = str_or(&value.get("author"), "Unknown");
    let date = st(&value.get("date")).filter(|date| !date.is_empty());
    let id = value.get("revisionId").or_else(|| value.get("id"));
    let info = obj! { "id": revision_id(&id), "author": author };
    if let (Some(date), Some(object)) = (date, info.obj()) {
        object.set("date", V::Str(date));
    }
    Some(info)
}

pub(crate) fn create_hyperlink(attributes: &V) -> Option<V> {
    let link = o(&attributes.get("hyperlink"))?;
    let href = st(&link.get("href"))
        .map(|href| href.to_string())
        .unwrap_or_default();
    if let Some(anchor) = href.strip_prefix('#') {
        return Some(obj! {
            "type": "hyperlink",
            "anchor": anchor,
            "tooltip": str_or_undef(&link.get("tooltip")),
            "children": arr![],
        });
    }
    Some(obj! {
        "type": "hyperlink",
        "href": href,
        "tooltip": str_or_undef(&link.get("tooltip")),
        "rId": str_or_undef(&link.get("rId")),
        "children": arr![],
    })
}

pub(crate) fn hyperlink_key(attributes: &V) -> Option<String> {
    let link = o(&attributes.get("hyperlink"))?;
    Some(
        st(&link.get("href"))
            .map(|href| href.to_string())
            .unwrap_or_default(),
    )
}

pub(crate) fn field_from_payload(payload: &V, attributes: &V) -> V {
    if let Some(field_data) = st(&payload.get("fieldData"))
        && !field_data.is_empty()
        && text::len(&field_data) <= 2_000_000
        && let Ok(stored) = V::parse(&field_data)
        && stored.obj().is_some()
    {
        let kind = ty(&stored);
        let children = if &*kind == "simpleField" {
            stored.get("content")
        } else {
            stored.get("fieldResult")
        };
        if (&*kind == "simpleField" || &*kind == "complexField")
            && stored.get("instruction").as_str().is_some()
            && children.arr().is_some()
        {
            if let Some(mode) = st(&payload.get("displayMode")).filter(|mode| !mode.is_empty())
                && let Some(tree) = stored
                    .get("fieldTree")
                    .obj()
                    .filter(|_| stored.get("fieldTree").truthy())
            {
                tree.set("displayMode", V::Str(mode));
            }
            return stored;
        }
    }
    let formatting = attrs_to_text_formatting(&formatting_attrs(attributes));
    let display = obj! {
        "type": "run",
        "content": arr![obj! { "type": "text", "text": str_or(&payload.get("displayText"), "") }],
    };
    if !formatting.is_empty()
        && let Some(run) = display.obj()
    {
        run.set("formatting", V::Obj(formatting));
    }
    let instruction = str_or(&payload.get("instruction"), "");
    let field_type = str_or(&payload.get("fieldType"), "UNKNOWN");
    let flag = |key: &str| {
        if payload.get(key).same(&V::Bool(true)) {
            V::Bool(true)
        } else {
            V::Undef
        }
    };
    let display_mode = str_or(&payload.get("displayMode"), "result");
    let tree = obj! { "version": 1.0, "displayMode": display_mode };
    if payload.get("fieldKind").as_str().as_deref() == Some("complex") {
        return obj! {
            "type": "complexField",
            "instruction": instruction,
            "fieldType": field_type,
            "fieldCode": arr![],
            "fieldResult": arr![display],
            "fldLock": flag("fldLock"),
            "dirty": flag("dirty"),
            "fieldTree": tree,
        };
    }
    obj! {
        "type": "simpleField",
        "instruction": instruction,
        "fieldType": field_type,
        "content": arr![display],
        "fldLock": flag("fldLock"),
        "dirty": flag("dirty"),
        "fieldTree": tree,
    }
}

fn math_from_payload(payload: &V) -> V {
    obj! {
        "type": "mathEquation",
        "display": str_or(&payload.get("display"), "inline"),
        "ommlXml": str_or(&payload.get("ommlXml"), ""),
        "plainText": str_or_undef(&payload.get("plainText")),
    }
}

fn horizontal_rule_run(payload: &V, attributes: &V) -> Result<V> {
    let value = payload.get("rule");
    let height = value.get("height").finite();
    let malformed = || Error::engine("Malformed horizontalRule embed payload");
    let Some(rule_object) = o(&value) else {
        return Err(malformed());
    };
    let width = rule_object.get("width");
    let percent = rule_object.get("widthPercent");
    if height.is_none()
        || (!width.nullish() && width.finite().is_none())
        || (!percent.nullish() && percent.finite().is_none())
        || rule_object.get("alignment").as_str().is_none()
        || rule_object.get("noShade").as_bool().is_none()
        || rule_object.get("color").as_str().is_none()
        || rule_object.get("xml").as_str().is_none()
    {
        return Err(malformed());
    }
    let rule = obj! {
        "width": width.finite().map(V::Num).unwrap_or(V::Null),
        "widthPercent": percent.finite().map(V::Num).unwrap_or(V::Null),
        "height": height,
        "alignment": rule_object.get("alignment"),
        "noShade": rule_object.get("noShade"),
        "color": rule_object.get("color"),
        "xml": rule_object.get("xml"),
    };
    let formatting = attrs_to_text_formatting(&formatting_attrs(attributes));
    let run = obj! {
        "type": "run",
        "content": arr![obj! { "type": "horizontalRule", "rule": rule }],
    };
    if !formatting.is_empty()
        && let Some(object) = run.obj()
    {
        object.set("formatting", V::Obj(formatting));
    }
    Ok(run)
}

/// The authored XML an unedited picture or shape replays.
fn with_source_xml(target: &V, payload: &V) {
    if let Some(xml) = st(&payload.get("sourceXml")).filter(|xml| !xml.is_empty())
        && let Some(object) = target.obj()
    {
        object.set("sourceXml", V::Str(xml));
    }
}

/// `text.match(/rotate\(([-\d.]+)deg\)/)?.[1]` parsed as `parseFloat`.
fn rotation_of(transform: &str) -> Option<f64> {
    let mut rest = transform;
    while let Some(at) = rest.find("rotate(") {
        let after = &rest[at + 7..];
        let digits: String = after
            .chars()
            .take_while(|ch| ch.is_ascii_digit() || *ch == '-' || *ch == '.')
            .collect();
        if !digits.is_empty() && after[digits.len()..].starts_with("deg)") {
            return Some(parse_float(&digits));
        }
        rest = &rest[at + 7..];
    }
    None
}

/// `Number.parseFloat` of a `[-\d.]+` string.
fn parse_float(text: &str) -> f64 {
    let mut end = 0;
    let mut seen_dot = false;
    let mut seen_digit = false;
    for (index, ch) in text.char_indices() {
        match ch {
            '-' if index == 0 => {}
            '.' if !seen_dot => seen_dot = true,
            '0'..='9' => seen_digit = true,
            _ => break,
        }
        end = index + ch.len_utf8();
    }
    if !seen_digit {
        return f64::NAN;
    }
    text[..end]
        .trim_end_matches('.')
        .parse()
        .unwrap_or(f64::NAN)
}

fn image_run_from_payload(payload: &V) -> Result<V> {
    let get = |key: &str| payload.get(key);
    let (wrap_value, wrap) = V::new_obj();
    wrap.set("type", str_or(&get("wrapType"), "inline"));
    for (from, to) in [
        ("distTop", "distT"),
        ("distBottom", "distB"),
        ("distLeft", "distL"),
        ("distRight", "distR"),
    ] {
        if !get(from).is_undef() {
            wrap.set(to, V::Num(px_to_emu(get(from).to_number())));
        }
    }
    if get("wrapText").truthy() {
        wrap.set("wrapText", get("wrapText"));
    }
    let number_or_zero = |value: V| {
        let number = value.to_number();
        if number.is_nan() || number == 0.0 {
            0.0
        } else {
            number
        }
    };
    let (image_value, image) = V::new_obj();
    image.set("type", V::str("image"));
    image.set("rId", str_or(&get("rId"), ""));
    image.set("src", str_or(&get("src"), ""));
    image.set("alt", str_or_undef(&get("alt")));
    image.set("title", str_or_undef(&get("title")));
    image.set("shapeType", str_or_undef(&get("shapeType")));
    image.set(
        "size",
        obj! {
            "width": px_to_emu(number_or_zero(get("width"))),
            "height": px_to_emu(number_or_zero(get("height"))),
        },
    );
    image.set("wrap", wrap_value);
    let transform = get("transform");
    if transform.truthy() {
        let Some(transform) = transform.as_str() else {
            return Err(Error::engine("attrs.transform.match is not a function"));
        };
        let (value, object) = V::new_obj();
        if let Some(rotation) = rotation_of(&transform) {
            object.set("rotation", V::Num(rotation));
        }
        if transform.contains("scaleX(-1)") {
            object.set("flipH", V::Bool(true));
        }
        if transform.contains("scaleY(-1)") {
            object.set("flipV", V::Bool(true));
        }
        if object.get("rotation").truthy()
            || object.get("flipH").truthy()
            || object.get("flipV").truthy()
        {
            image.set("transform", value);
        }
    }
    let position = get("position");
    let (horizontal, vertical) = (position.get("horizontal"), position.get("vertical"));
    if horizontal.truthy() && vertical.truthy() {
        image.set(
            "position",
            obj! {
                "relativeHeight": position.get("relativeHeight"),
                "horizontal": obj! {
                    "relativeTo": horizontal.get("relativeTo").or(|| V::str("column")),
                    "alignment": horizontal.get("align"),
                    "posOffset": horizontal.get("posOffset"),
                },
                "vertical": obj! {
                    "relativeTo": vertical.get("relativeTo").or(|| V::str("paragraph")),
                    "alignment": vertical.get("align"),
                    "posOffset": vertical.get("posOffset"),
                },
            },
        );
    }
    let border = get("borderWidth");
    if border.truthy() && border.to_number() > 0.0 {
        let color = get("borderColor");
        let style = get("borderStyle").as_str().and_then(|style| match &*style {
            "solid" | "double" | "groove" | "ridge" | "inset" | "outset" => Some("solid"),
            "dotted" => Some("dot"),
            "dashed" => Some("dash"),
            _ => None,
        });
        image.set(
            "outline",
            obj! {
                "width": px_to_emu(border.to_number()),
                "color": if color.truthy() {
                    obj! { "rgb": color.to_js_string().replacen('#', "", 1) }
                } else {
                    V::Undef
                },
                "style": style.unwrap_or("solid"),
            },
        );
    }
    if get("hlinkHref").truthy() {
        image.set("hlinkHref", get("hlinkHref"));
    }
    let (crop_value, crop) = V::new_obj();
    for (from, to) in [
        ("cropTop", "top"),
        ("cropRight", "right"),
        ("cropBottom", "bottom"),
        ("cropLeft", "left"),
    ] {
        if !get(from).nullish() {
            crop.set(to, get(from));
        }
    }
    if !crop.is_empty() {
        image.set("crop", crop_value);
    }
    let opacity = get("opacity");
    if !opacity.nullish() && opacity.to_number() < 1.0 {
        image.set("opacity", opacity);
    }
    for key in ["layoutInCell", "allowOverlap"] {
        if !get(key).nullish() {
            image.set(key, get(key));
        }
    }
    let (padding_value, padding) = V::new_obj();
    for (from, to) in [
        ("effectExtentTop", "top"),
        ("effectExtentBottom", "bottom"),
        ("effectExtentLeft", "left"),
        ("effectExtentRight", "right"),
    ] {
        if get(from).truthy() {
            padding.set(to, V::Num(px_to_emu(get(from).to_number())));
        }
    }
    if !padding.is_empty() {
        image.set("padding", padding_value);
    }
    let drawing = obj! { "type": "drawing", "image": image_value };
    with_source_xml(&drawing, payload);
    Ok(obj! { "type": "run", "content": arr![drawing] })
}

/// The seeded shape, carrying the text body and everything else no payload field describes.
fn stored_shape(value: &V) -> Option<V> {
    let json = st(value).filter(|json| !json.is_empty())?;
    if text::len(&json) > 2_000_000 {
        return None;
    }
    let parsed = V::parse(&json).ok()?;
    if !is(&parsed, "shape") || parsed.get("shapeType").as_str().is_none() {
        return None;
    }
    if parsed.get("size").obj().is_none()
        && let Some(object) = parsed.obj()
    {
        object.set("size", obj! { "width": 0.0, "height": 0.0 });
    }
    Some(parsed)
}

fn chart_run_from_payload(payload: &V) -> Option<V> {
    let json = st(&payload.get("chartJson")).filter(|json| !json.is_empty())?;
    let chart = V::parse(&json).ok()?;
    if !is(&chart, "chart") || chart.get("chartType").as_str().is_none() {
        return None;
    }
    Some(obj! { "type": "run", "content": arr![obj! { "type": "chart", "chart": chart }] })
}

fn opaque_drawing_run(payload: &V) -> V {
    obj! {
        "type": "run",
        "content": arr![obj! {
            "type": "opaqueDrawing",
            "kind": st(&payload.get("kind")).map(V::Str).unwrap_or(V::str("")),
            "xml": st(&payload.get("xml")).map(V::Str).unwrap_or(V::str("")),
        }],
    }
}

fn hex_rgb(value: &str) -> String {
    value.replacen('#', "", 1)
}

fn shape_run_from_payload(payload: &V) -> V {
    let get = |key: &str| payload.get(key);
    let shape = stored_shape(&get("shapeJson")).unwrap_or_else(|| {
        obj! {
            "type": "shape",
            "shapeType": "rect",
            "size": obj! { "width": 0.0, "height": 0.0 },
        }
    });
    let object = o(&shape).unwrap_or_default();
    if let Some(kind) = st(&get("shapeType")).filter(|kind| !kind.is_empty()) {
        object.set("shapeType", V::Str(kind));
    }
    if let Some(id) = st(&get("shapeId")).filter(|id| !id.is_empty()) {
        object.set("id", V::Str(id));
    }
    for key in ["width", "height"] {
        if get(key).truthy() {
            let size = o(&object.get("size"))
                .map(|size| size.spread())
                .unwrap_or_default();
            size.set(key, V::Num(px_to_emu(get(key).to_number())));
            object.set("size", V::Obj(size));
        }
    }
    if let Some(path) = get("geometryPath").arr().filter(|path| !path.is_empty()) {
        object.set("geometryPath", V::Arr(path.clone()));
    }
    let fill_type = get("fillType");
    if fill_type.as_str().as_deref() == Some("gradient") && get("gradientStops").as_str().is_some()
    {
        let stops = V::parse(&get("gradientStops").to_js_string())
            .ok()
            .and_then(|stops| {
                let items = stops.arr()?.items();
                items
                    .iter()
                    .map(|stop| {
                        let color = stop.get("color").as_str()?;
                        Some(obj! {
                            "position": stop.get("position"),
                            "color": obj! { "rgb": hex_rgb(&color) },
                        })
                    })
                    .collect::<Option<Vec<_>>>()
                    .filter(|_| stops.arr().is_some())
            });
        match stops {
            Some(stops) => {
                let angle = get("gradientAngle").finite().filter(|angle| *angle != 0.0);
                object.set(
                    "fill",
                    obj! {
                        "type": "gradient",
                        "gradient": obj! {
                            "type": str_or(&get("gradientType"), "linear"),
                            "angle": angle,
                            "stops": V::array(stops),
                        },
                    },
                );
            }
            None => {
                let color = st(&get("fillColor"))
                    .filter(|color| !color.is_empty())
                    .map(|color| color.to_string())
                    .unwrap_or_else(|| "000000".into());
                object.set(
                    "fill",
                    obj! { "type": "solid", "color": obj! { "rgb": hex_rgb(&color) } },
                );
            }
        }
    } else if let Some(color) = get("fillColor").as_str() {
        object.set(
            "fill",
            obj! {
                "type": str_or(&fill_type, "solid"),
                "color": obj! { "rgb": hex_rgb(&color) },
            },
        );
    } else if fill_type.as_str().as_deref() == Some("none") {
        object.set("fill", obj! { "type": "none" });
    }
    if let Some(width) = get("outlineWidth").as_num().filter(|width| *width > 0.0) {
        let style = get("outlineStyle")
            .as_str()
            .and_then(|style| match &*style {
                "solid" => Some("solid"),
                "dotted" => Some("dot"),
                "dashed" => Some("dash"),
                _ => None,
            });
        object.set(
            "outline",
            obj! {
                "width": px_to_emu(width),
                "color": get("outlineColor").as_str().map(|color| obj! { "rgb": hex_rgb(&color) }),
                "style": style.unwrap_or("solid"),
            },
        );
    }
    let (transform_value, transform) = V::new_obj();
    if let Some(rotation) = get("rotation").as_num() {
        transform.set("rotation", V::Num(rotation));
    } else if let Some(text) = get("transform").as_str()
        && let Some(rotation) = rotation_of(&text)
    {
        transform.set("rotation", V::Num(rotation));
    }
    let transform_text = get("transform").or(|| V::str("")).to_js_string();
    if get("flipH").truthy() || transform_text.contains("scaleX(-1)") {
        transform.set("flipH", V::Bool(true));
    }
    if get("flipV").truthy() || transform_text.contains("scaleY(-1)") {
        transform.set("flipV", V::Bool(true));
    }
    if transform.get("rotation").truthy()
        || transform.get("flipH").truthy()
        || transform.get("flipV").truthy()
    {
        object.set("transform", transform_value);
    }
    let entry = obj! { "type": "shape", "shape": shape };
    with_source_xml(&entry, payload);
    obj! { "type": "run", "content": arr![entry] }
}

pub(crate) fn inline_sdt_from_payload(payload: &V) -> Result<V> {
    let mut properties = sdt_attrs_to_props(payload);
    if let Some(json) = st(&payload.get("propertiesJson")).filter(|json| !json.is_empty())
        && text::len(&json) <= 1_000_000
        && let Ok(parsed) = V::parse(&json)
        && parsed.obj().is_some()
        && parsed.get("sdtType").as_str().is_some()
    {
        properties = parsed;
    }
    let mut items = Vec::new();
    for raw in payload.get("content").items() {
        let Some(entry) = o(&raw) else { continue };
        let Some(kind) = st(&entry.get("kind")).filter(|kind| !kind.is_empty()) else {
            continue;
        };
        let attributes = o(&entry.get("attrs"))
            .map(V::Obj)
            .unwrap_or_else(|| obj! {});
        if &*kind == "text" {
            if let Some(text) = st(&entry.get("text")) {
                items.push(obj! { "kind": "text", "text": V::Str(text), "attributes": attributes });
            }
            continue;
        }
        items.push(obj! {
            "kind": "embed",
            "embedKind": V::Str(kind),
            "payload": o(&entry.get("payload")).map(V::Obj).unwrap_or_else(|| obj! {}),
            "attributes": attributes,
        });
    }
    let mut content = inline_sdt_content(&build_paragraph_content(items)?);
    if let Some(value) = content_control_value(&payload.get("value"))
        && let Ok((applied, blocks)) = apply_content_control_value(&properties, &value)
    {
        properties = applied;
        let display = blocks.first().cloned().unwrap_or_default();
        content = inline_sdt_content(&if is(&display, "paragraph") {
            display.get("content").items()
        } else {
            Vec::new()
        });
    }
    Ok(obj! { "type": "inlineSdt", "properties": properties, "content": V::array(content) })
}

pub(crate) fn inline_sdt_content(content: &[V]) -> Vec<V> {
    content
        .iter()
        .filter(|child| {
            !matches!(
                &*ty(child),
                "commentRangeStart"
                    | "commentRangeEnd"
                    | "moveFromRangeStart"
                    | "moveFromRangeEnd"
                    | "moveToRangeStart"
                    | "moveToRangeEnd"
            )
        })
        .cloned()
        .collect()
}

pub(crate) fn content_control_value(value: &V) -> Option<V> {
    let authored = o(value)?;
    let kind = authored.get("kind").as_str();
    match kind.as_deref() {
        Some("checkbox") if authored.get("checked").as_bool().is_some() => {
            Some(obj! { "kind": "checkbox", "checked": authored.get("checked") })
        }
        Some("dropdown") if authored.get("value").as_str().is_some() => {
            Some(obj! { "kind": "dropdown", "value": authored.get("value") })
        }
        Some("date") if authored.get("date").as_str().is_some() => {
            Some(obj! { "kind": "date", "date": authored.get("date") })
        }
        _ => None,
    }
}

fn comment_reference_from_payload(payload: &V) -> Option<V> {
    if payload.get("modelKind").as_str().as_deref() != Some("commentReference") {
        return None;
    }
    let reference = obj! { "type": "commentReference" };
    if let Some(id) = payload.get("commentId").finite()
        && let Some(object) = reference.obj()
    {
        object.set("id", V::Num(id));
    }
    Some(obj! { "type": "run", "content": arr![reference] })
}

/// Page and column break runs; [`Breaks`] remembers which are story units.
pub(crate) fn break_run(kind: &str) -> V {
    obj! { "type": "run", "content": arr![obj! { "type": "break", "breakType": kind }] }
}

thread_local! {
    /// yrsToDocument.ts `unitBreaks`: break entries that are story units,
    /// kept alive so their identities stay unique.
    static UNIT_BREAKS: std::cell::RefCell<HashMap<usize, Obj>> =
        std::cell::RefCell::new(HashMap::new());
}

pub(crate) fn unit_break_run(kind: &str) -> V {
    let run = break_run(kind);
    if let Some(entry) = run.get("content").arr().map(|content| content.get(0))
        && let Some(entry) = entry.obj()
    {
        UNIT_BREAKS.with(|breaks| breaks.borrow_mut().insert(entry.id(), entry.clone()));
    }
    run
}

pub(crate) fn is_unit_break(entry: &V) -> bool {
    entry
        .obj()
        .is_some_and(|entry| UNIT_BREAKS.with(|breaks| breaks.borrow().contains_key(&entry.id())))
}

/// Forgets the unit breaks of a finished projection.
pub(crate) fn clear_unit_breaks() {
    UNIT_BREAKS.with(|breaks| breaks.borrow_mut().clear());
}

pub(crate) fn raw_content_for_item(item: &V) -> Result<Option<V>> {
    let attributes = item.get("attributes");
    if is_kind(item, "text") {
        return Ok(Some(create_text_run(
            &item.get("text").to_js_string(),
            &attributes,
        )));
    }
    let payload = item.get("payload");
    let embed = item
        .get("embedKind")
        .as_str()
        .unwrap_or_else(|| Str::from(""));
    Ok(match &*embed {
        "bookmark" => {
            let kind = payload.get("kind").as_str();
            if matches!(kind.as_deref(), Some("fieldseparate" | "fieldend")) {
                let run = payload.get("run");
                (is(&run, "run") && run.get("content").arr().is_some()).then_some(run)
            } else {
                let boundary = o(&payload)
                    .map(|payload| payload.spread())
                    .unwrap_or_default();
                boundary.set("offset", V::Num(0.0));
                Some(super::paragraph::bookmark_node(&V::Obj(boundary)))
            }
        }
        "break" => Some(obj! {
            "type": "run",
            "content": arr![obj! { "type": "break", "breakType": "textWrapping" }],
        }),
        "flowBreak" => Some(unit_break_run(&payload.get("breakType").to_js_string())),
        "pageBreak" => Some(break_run("page")),
        "columnBreak" => Some(break_run("column")),
        "tab" => Some(match o(&payload.get("ptab")) {
            // A positional tab keeps its run's formatting, as text does.
            Some(ptab) => {
                let run = o(&create_text_run("", &attributes)).unwrap_or_default();
                let entry = ptab.spread();
                entry.set("type", V::str("ptab"));
                run.set("content", arr![V::Obj(entry)]);
                V::Obj(run)
            }
            None => obj! { "type": "run", "content": arr![obj! { "type": "tab" }] },
        }),
        "image" => Some(image_run_from_payload(&payload)?),
        "horizontalRule" => Some(horizontal_rule_run(&payload, &attributes)?),
        "shape" => Some(shape_run_from_payload(&payload)),
        "chart" => chart_run_from_payload(&payload),
        "opaqueDrawing" => Some(opaque_drawing_run(&payload)),
        "field" => Some(
            comment_reference_from_payload(&payload)
                .unwrap_or_else(|| field_from_payload(&payload, &attributes)),
        ),
        "math" => Some(math_from_payload(&payload)),
        "sdt" => Some(inline_sdt_from_payload(&payload)?),
        "noteRef" => {
            let footnote = payload.get("footnoteRefId");
            let endnote = payload.get("endnoteRefId");
            // The reference keeps its run's formatting (its FootnoteReference style).
            let run = o(&create_text_run("", &attributes)).unwrap_or_default();
            run.set(
                "content",
                arr![if !footnote.is_undef() {
                    obj! { "type": "footnoteRef", "id": revision_id(&footnote) }
                } else {
                    obj! { "type": "endnoteRef", "id": revision_id(&endnote) }
                }],
            );
            Some(V::Obj(run))
        }
        _ => None,
    })
}

pub(crate) fn is_kind(item: &V, kind: &str) -> bool {
    item.get("kind").as_str().as_deref() == Some(kind)
}

pub(crate) fn ordinary_content_for_item(item: &V) -> Result<Option<V>> {
    let content = raw_content_for_item(item)?;
    let changes = run_property_changes(&item.get("attributes"));
    Ok(match content {
        Some(content) if is(&content, "run") && changes.truthy() => Some(V::Obj(
            o(&content)
                .unwrap_or_default()
                .with(&[("propertyChanges", changes)]),
        )),
        other => other,
    })
}

pub(crate) fn tracked_content_for_item(item: &V, info: &V, linked: bool) -> Result<V> {
    let attributes = item.get("attributes");
    let link = if linked {
        create_hyperlink(&attributes)
    } else {
        None
    };
    if let Some(link) = &link {
        add_to_hyperlink(link, item)?;
    }
    let child = match link {
        Some(link) => link,
        None => ordinary_content_for_item(item)?
            .unwrap_or_else(|| obj! { "type": "run", "content": arr![] }),
    };
    let mut content = V::array(inline_sdt_content(&[child]));
    let mut result = V::Undef;
    for (key, normal, moved) in [
        ("del", "deletion", "moveFrom"),
        ("ins", "insertion", "moveTo"),
    ] {
        let raw = attributes.get(key);
        let Some(raw_object) = o(&raw) else { continue };
        let kind = if raw_object.get("isMovePair").same(&V::Bool(true)) {
            moved
        } else {
            normal
        };
        result = obj! {
            "type": kind,
            "info": tracked_info(&raw).unwrap_or_else(|| info.clone()),
            "content": content,
        };
        content = arr![result.clone()];
    }
    Ok(result)
}

pub(crate) fn add_to_hyperlink(hyperlink: &V, item: &V) -> Result<()> {
    let Some(child) = ordinary_content_for_item(item)? else {
        return Ok(());
    };
    if is(&child, "hyperlink") {
        return Ok(());
    }
    let Some(inline) = inline_sdt_content(&[child]).into_iter().next() else {
        return Ok(());
    };
    if is(&inline, "rawXml") || is(&inline, "hyperlink") {
        return Ok(());
    }
    let Some(link) = hyperlink.obj() else {
        return Ok(());
    };
    if !is(&inline, "run") && link.get("structuredChildren").nullish() {
        link.set(
            "structuredChildren",
            V::Arr(
                link.get("children")
                    .arr()
                    .map(Arr::copy)
                    .unwrap_or_default(),
            ),
        );
    }
    if let Some(children) = link.get("children").arr() {
        for run in shown_runs(std::slice::from_ref(&inline)) {
            children.push(run);
        }
    }
    if let Some(structured) = link.get("structuredChildren").arr() {
        structured.push(inline);
    }
    Ok(())
}

fn projection_signature(items: &[V]) -> String {
    let mut normalized: Vec<V> = Vec::new();
    for source in items {
        let attributes = o(&source.get("attributes"))
            .map(|object| object.spread())
            .unwrap_or_default();
        attributes.delete("fieldResult");
        let item = if is_kind(source, "embed")
            && source.get("embedKind").as_str().as_deref() == Some("tab")
        {
            obj! { "kind": "text", "text": "\t", "attributes": V::Obj(attributes.clone()) }
        } else {
            V::Obj(
                o(source)
                    .unwrap_or_default()
                    .with(&[("attributes", V::Obj(attributes.clone()))]),
            )
        };
        if let Some(previous) = normalized.last()
            && is_kind(&item, "text")
            && is_kind(previous, "text")
            && stable_stringify(&previous.get("attributes"))
                == stable_stringify(&V::Obj(attributes))
        {
            let joined = format!(
                "{}{}",
                previous.get("text").to_js_string(),
                item.get("text").to_js_string()
            );
            if let Some(previous) = previous.obj() {
                previous.set("text", V::from(joined));
            }
        } else {
            normalized.push(item);
        }
    }
    stable_stringify(&V::array(normalized))
}

/// A field's projected children by their index in its result.
type ChildGroups = Vec<(f64, Vec<V>)>;

/// Moves each field's projected children back into its result
/// (`restoreProjectedFieldResults`).
pub(crate) fn restore_projected_field_results(items: Vec<V>) -> Result<Vec<V>> {
    let mut owners: Vec<V> = Vec::new();
    let mut owner_at: HashMap<usize, V> = HashMap::new();
    let mut run = 0_usize;
    for (position, item) in items.iter().enumerate() {
        if item.get("attributes").get("fieldResult").truthy() {
            run += 1;
            continue;
        }
        let projection = if is_kind(item, "embed") {
            o(&item.get("payload").get("resultProjection"))
        } else {
            None
        };
        let id = projection
            .as_ref()
            .and_then(|projection| projection.get("id").finite());
        if is_kind(item, "embed")
            && let Some(id) = id
        {
            owners.push(item.clone());
            let recorded: Vec<Option<f64>> = projection
                .as_ref()
                .map(|projection| projection.get("children").items())
                .unwrap_or_default()
                .iter()
                .map(|child| child.get("index").finite())
                .collect();
            let mut at = position as isize - 1;
            while at >= position as isize - run as isize {
                let marker = items[at as usize].get("attributes").get("fieldResult");
                let marker = o(&marker);
                let marker_id = marker.as_ref().and_then(|marker| marker.get("id").finite());
                let marker_index = marker
                    .as_ref()
                    .and_then(|marker| marker.get("index").finite());
                if marker_id != Some(id) || !recorded.contains(&marker_index) {
                    break;
                }
                owner_at.insert(at as usize, item.clone());
                at -= 1;
            }
        }
        run = 0;
    }
    if owners.is_empty() {
        return Ok(items);
    }
    // Groups by owner identity, then by child index, in insertion order.
    let mut groups: Vec<(V, ChildGroups)> = Vec::new();
    let mut remaining = Vec::new();
    for (position, item) in items.iter().enumerate() {
        let marker = item.get("attributes").get("fieldResult");
        if !marker.truthy() {
            remaining.push(item.clone());
            continue;
        }
        let index = marker.get("index").finite();
        let owner = owner_at.get(&position);
        let attributes = o(&item.get("attributes"))
            .map(|object| object.spread())
            .unwrap_or_default();
        attributes.delete("fieldResult");
        let copy = V::Obj(
            o(item)
                .unwrap_or_default()
                .with(&[("attributes", V::Obj(attributes))]),
        );
        let (Some(index), Some(owner)) = (index, owner) else {
            remaining.push(copy);
            continue;
        };
        let slot = match groups.iter().position(|(key, _)| key.same(owner)) {
            Some(slot) => slot,
            None => {
                groups.push((owner.clone(), Vec::new()));
                groups.len() - 1
            }
        };
        let children = &mut groups[slot].1;
        match children.iter_mut().find(|(key, _)| *key == index) {
            Some((_, group)) => group.push(copy),
            None => children.push((index, vec![copy])),
        }
    }
    for owner in &owners {
        let stored = field_from_payload(&owner.get("payload"), &owner.get("attributes"));
        if !is(&stored, "complexField") {
            continue;
        }
        let projection = owner.get("payload").get("resultProjection");
        let originals = projection.get("children").items();
        let mut replacements: Vec<(f64, Vec<V>)> = Vec::new();
        for raw in &originals {
            let Some(child) = o(raw) else { continue };
            let Some(index) = child.get("index").finite() else {
                continue;
            };
            let Some(original_items) = child.get("items").arr().cloned() else {
                continue;
            };
            let current: Vec<V> = groups
                .iter()
                .find(|(key, _)| key.same(owner))
                .and_then(|(_, children)| children.iter().find(|(key, _)| *key == index))
                .map(|(_, group)| group.clone())
                .unwrap_or_default();
            if projection_signature(&current) == projection_signature(&original_items.items()) {
                continue;
            }
            let mut rebuilt = inline_sdt_content(&build_paragraph_content(current)?);
            let original = if index < 0.0 {
                stored
                    .get("structuredCode")
                    .get("inline")
                    .arr()
                    .map(|inline| inline.get((-index - 1.0) as usize))
                    .unwrap_or_default()
            } else {
                stored
                    .get("structuredResult")
                    .get("inline")
                    .arr()
                    .map(|inline| inline.get(index as usize))
                    .unwrap_or_default()
            };
            let links: Vec<usize> = rebuilt
                .iter()
                .enumerate()
                .filter(|(_, node)| is(node, "hyperlink"))
                .map(|(at, _)| at)
                .collect();
            if is(&original, "hyperlink") && links.len() == 1 {
                let link = rebuilt[links[0]].clone();
                let merged = o(&original).unwrap_or_default().spread();
                merged.assign(&o(&link).unwrap_or_default());
                merged.set("structuredChildren", link.get("structuredChildren"));
                rebuilt[links[0]] = V::Obj(merged);
            }
            replacements.push((index, rebuilt));
        }
        if replacements.is_empty() {
            continue;
        }
        let replaced = |list: &V, sign: f64| -> Option<Vec<V>> {
            let items = list.arr()?.items();
            Some(
                items
                    .iter()
                    .enumerate()
                    .flat_map(|(at, child)| {
                        let key = if sign < 0.0 {
                            -(at as f64) - 1.0
                        } else {
                            at as f64
                        };
                        replacements
                            .iter()
                            .find(|(index, _)| *index == key)
                            .map(|(_, rebuilt)| rebuilt.clone())
                            .unwrap_or_else(|| vec![child.clone()])
                    })
                    .collect(),
            )
        };
        let stored_object = o(&stored).unwrap_or_default();
        let inline = V::array(
            replaced(&stored.get("structuredResult").get("inline"), 1.0).unwrap_or_default(),
        );
        let code = replaced(&stored.get("structuredCode").get("inline"), -1.0).map(V::array);
        let result = o(&stored.get("structuredResult"))
            .map(|object| object.spread())
            .unwrap_or_default();
        result.set("inline", inline.clone());
        stored_object.set("structuredResult", V::Obj(result));
        if let Some(code) = &code {
            let structured = o(&stored.get("structuredCode"))
                .map(|object| object.spread())
                .unwrap_or_default();
            structured.set("inline", code.clone());
            stored_object.set("structuredCode", V::Obj(structured));
        }
        let tree = stored.get("fieldTree");
        if tree.truthy()
            && let Some(tree) = tree.obj()
        {
            let result = o(&tree.get("result"))
                .map(|object| object.spread())
                .unwrap_or_default();
            result.set("inline", inline.clone());
            tree.set("result", V::Obj(result));
            if let Some(code) = &code {
                let code_tree = o(&tree.get("code"))
                    .map(|object| object.spread())
                    .unwrap_or_default();
                code_tree.set("inline", code.clone());
                tree.set("code", V::Obj(code_tree));
            }
        }
        stored_object.set("fieldResult", V::array(shown_runs(&inline.items())));
        if let Some(owner_object) = owner.obj() {
            let payload = o(&owner.get("payload"))
                .unwrap_or_default()
                .with(&[("fieldData", V::from(stringify(&stored).unwrap_or_default()))]);
            owner_object.set("payload", V::Obj(payload));
        }
    }
    Ok(remaining)
}

/// The runs field result `nodes` show, as the parser reads them (`shownRuns`).
pub(crate) fn shown_runs(nodes: &[V]) -> Vec<V> {
    nodes
        .iter()
        .flat_map(|node| match &*ty(node) {
            "run" => vec![node.clone()],
            "hyperlink" => shown_runs(&node.get("children").items()),
            "simpleField" => shown_runs(&node.get("content").items()),
            "complexField" => node.get("fieldResult").items(),
            "inlineSdt" => shown_runs(&node.get("content").items()),
            "insertion" | "moveTo" => shown_runs(&node.get("content").items()),
            "rawXml" => node.get("shown").items(),
            _ => Vec::new(),
        })
        .collect()
}

pub(crate) fn build_paragraph_content(items: Vec<V>) -> Result<Vec<V>> {
    let items = restore_projected_field_results(items)?;
    let mut content: Vec<V> = Vec::new();
    let mut current_run: Option<V> = None;
    let mut current_key: Option<String> = None;
    let mut current_link: Option<V> = None;
    fn flush_run(content: &mut Vec<V>, run: &mut Option<V>, key: &mut Option<String>) {
        if let Some(run) = run.take() {
            content.push(run);
        }
        *key = None;
    }
    fn flush_link(content: &mut Vec<V>, link: &mut Option<V>) {
        if let Some(link) = link.take() {
            content.push(link);
        }
    }
    for item in &items {
        let attributes = item.get("attributes");
        // A note reference is handled before link marks; a tracked one keeps its revision.
        if is_kind(item, "embed") && item.get("embedKind").as_str().as_deref() == Some("noteRef") {
            flush_run(&mut content, &mut current_run, &mut current_key);
            flush_link(&mut content, &mut current_link);
            let revision = tracked_info(&attributes.get("ins").or_else(|| attributes.get("del")));
            let note = match revision {
                Some(revision) => Some(tracked_content_for_item(item, &revision, false)?),
                None => ordinary_content_for_item(item)?,
            };
            if let Some(note) = note {
                content.push(note);
            }
            continue;
        }
        if let Some(revision) =
            tracked_info(&attributes.get("ins").or_else(|| attributes.get("del")))
        {
            flush_run(&mut content, &mut current_run, &mut current_key);
            flush_link(&mut content, &mut current_link);
            content.push(tracked_content_for_item(item, &revision, true)?);
            continue;
        }
        if let Some(link_key) = hyperlink_key(&attributes) {
            flush_run(&mut content, &mut current_run, &mut current_key);
            let current = current_link.as_ref().map(|link| {
                let href = link.get("href");
                if href.truthy() {
                    href.to_js_string()
                } else if link.get("anchor").truthy() {
                    format!("#{}", link.get("anchor").to_js_string())
                } else {
                    String::new()
                }
            });
            if current.as_deref() != Some(link_key.as_str()) {
                flush_link(&mut content, &mut current_link);
                current_link = create_hyperlink(&attributes);
            }
            if let Some(link) = &current_link {
                add_to_hyperlink(link, item)?;
            }
            continue;
        }
        flush_link(&mut content, &mut current_link);
        if is_kind(item, "text") {
            let key = stable_stringify(&formatting_attrs(&attributes));
            let next = create_text_run(&item.get("text").to_js_string(), &attributes);
            match (&current_run, &current_key) {
                (Some(run), Some(current)) if *current == key => append_text_run(run, &next),
                _ => {
                    flush_run(&mut content, &mut current_run, &mut current_key);
                    current_run = Some(next);
                    current_key = Some(key);
                }
            }
            continue;
        }
        flush_run(&mut content, &mut current_run, &mut current_key);
        if let Some(child) = ordinary_content_for_item(item)? {
            content.push(child);
        }
    }
    flush_run(&mut content, &mut current_run, &mut current_key);
    flush_link(&mut content, &mut current_link);
    Ok(content)
}
