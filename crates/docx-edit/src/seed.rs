use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::fmt;
use std::sync::Arc;

use serde::Deserialize;
use serde::de::{MapAccess, SeqAccess, Visitor};
use serde_json::{Map, Value, json};
use yrs::Any;
use yrs::types::Attrs;

use crate::{EditCtx, EditingDoc, RawOp};

pub(crate) type JsonObject = BTreeMap<String, Value>;

#[derive(Clone)]
struct Mark {
    name: String,
    attrs: Vec<(String, Value)>,
}

#[derive(Clone)]
enum UnitContent {
    Text(String),
    Embed { kind: String, payload: JsonObject },
}

#[derive(Clone)]
struct InlineUnit {
    content: UnitContent,
    attrs: JsonObject,
    pm_size: u32,
    /// Comments whose range holds this unit, in the order they opened.
    comments: Vec<String>,
    marks: Vec<Mark>,
}

struct StoryPlan {
    story_id: String,
    units: Vec<InlineUnit>,
    comment_coverage: Vec<(String, Vec<(u32, u32)>)>,
}

struct ProjectedCell {
    paragraph_formatting: Option<Value>,
    attrs: JsonObject,
    content: Vec<Value>,
}

struct ProjectedRow {
    attrs: JsonObject,
    cells: Vec<ProjectedCell>,
}

struct ProjectedTable {
    attrs: JsonObject,
    rows: Vec<ProjectedRow>,
}

#[derive(Clone, Copy)]
struct StoryOptions {
    append_body_tail: bool,
}

struct LoweringContext {
    styles: StyleResolver,
    theme: Option<Value>,
    source_json: Arc<BTreeMap<String, String>>,
    plans: Vec<StoryPlan>,
    compatibility_mode: u8,
    open_comments: OpenComments,
}

impl LoweringContext {
    /// Appends `unit` to plan `story`, inside the comments open there.
    fn push(&mut self, story: usize, mut unit: InlineUnit) {
        unit.comments = self.open_comments.covering(story);
        self.plans[story].units.push(unit);
    }
}

fn compatibility_mode_from_package(package: Option<&Value>) -> u8 {
    field(field(package, "settings"), "compatibilityFlags")
        .and_then(|flags| field(Some(flags), "compatibilityMode"))
        .and_then(|value| value.as_f64())
        .filter(|value| value.is_finite() && (0.0..=255.0).contains(value))
        .map(|value| value as u8)
        .unwrap_or(12)
}

enum OrderedValue {
    Null,
    Bool(bool),
    Number(serde_json::Number),
    String(String),
    Array(Vec<OrderedValue>),
    Object(Vec<(String, OrderedValue)>),
}

struct OrderedValueVisitor;

impl<'de> Visitor<'de> for OrderedValueVisitor {
    type Value = OrderedValue;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a JSON value")
    }

    fn visit_unit<E>(self) -> Result<Self::Value, E> {
        Ok(OrderedValue::Null)
    }

    fn visit_none<E>(self) -> Result<Self::Value, E> {
        Ok(OrderedValue::Null)
    }

    fn visit_bool<E>(self, value: bool) -> Result<Self::Value, E> {
        Ok(OrderedValue::Bool(value))
    }

    fn visit_i64<E>(self, value: i64) -> Result<Self::Value, E> {
        Ok(OrderedValue::Number(value.into()))
    }

    fn visit_u64<E>(self, value: u64) -> Result<Self::Value, E> {
        Ok(OrderedValue::Number(value.into()))
    }

    fn visit_f64<E>(self, value: f64) -> Result<Self::Value, E>
    where
        E: serde::de::Error,
    {
        serde_json::Number::from_f64(value)
            .map(OrderedValue::Number)
            .ok_or_else(|| E::custom("non-finite JSON number"))
    }

    fn visit_str<E>(self, value: &str) -> Result<Self::Value, E> {
        Ok(OrderedValue::String(value.to_owned()))
    }

    fn visit_string<E>(self, value: String) -> Result<Self::Value, E> {
        Ok(OrderedValue::String(value))
    }

    fn visit_seq<A>(self, mut sequence: A) -> Result<Self::Value, A::Error>
    where
        A: SeqAccess<'de>,
    {
        let mut values = Vec::new();
        while let Some(value) = sequence.next_element()? {
            values.push(value);
        }
        Ok(OrderedValue::Array(values))
    }

    fn visit_map<A>(self, mut object: A) -> Result<Self::Value, A::Error>
    where
        A: MapAccess<'de>,
    {
        let mut entries = Vec::new();
        while let Some(entry) = object.next_entry()? {
            entries.push(entry);
        }
        Ok(OrderedValue::Object(entries))
    }
}

impl<'de> Deserialize<'de> for OrderedValue {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        deserializer.deserialize_any(OrderedValueVisitor)
    }
}

#[derive(Default)]
struct StyleResolver {
    enabled: bool,
    styles: BTreeMap<String, Value>,
    doc_defaults: Option<Value>,
    default_paragraph: Option<String>,
    default_table: Option<String>,
    default_character: Option<String>,
    table_paragraph_formatting: Option<Value>,
}

fn object(value: Option<&Value>) -> Option<&Map<String, Value>> {
    value?.as_object()
}

fn array(value: Option<&Value>) -> &[Value] {
    value
        .and_then(Value::as_array)
        .map(Vec::as_slice)
        .unwrap_or(&[])
}

fn field<'a>(value: Option<&'a Value>, key: &str) -> Option<&'a Value> {
    object(value)?.get(key)
}

fn string(value: Option<&Value>) -> Option<&str> {
    value?.as_str()
}

fn number(value: Option<&Value>) -> Option<f64> {
    value?.as_f64()
}

fn boolean(value: Option<&Value>) -> Option<bool> {
    value?.as_bool()
}

fn truthy(value: Option<&Value>) -> bool {
    match value {
        None | Some(Value::Null) => false,
        Some(Value::Bool(value)) => *value,
        Some(Value::Number(value)) => value.as_f64().is_some_and(|value| value != 0.0),
        Some(Value::String(value)) => !value.is_empty(),
        Some(Value::Array(_) | Value::Object(_)) => true,
    }
}

fn nullish(value: Option<&Value>) -> Value {
    value.cloned().unwrap_or(Value::Null)
}

fn js_string(value: &Value) -> String {
    match value {
        Value::String(value) => value.clone(),
        Value::Number(value) => {
            if let Some(number) = value.as_f64()
                && number.fract() == 0.0
                && number.abs() <= 9_007_199_254_740_991.0
            {
                return format!("{number:.0}");
            }
            value.to_string()
        }
        Value::Bool(value) => value.to_string(),
        Value::Null => "null".to_owned(),
        _ => serde_json::to_string(value).unwrap_or_default(),
    }
}

fn utf16_len(value: &str) -> u32 {
    value.encode_utf16().count() as u32
}

fn ordered_object(
    entries: impl IntoIterator<Item = (impl Into<String>, Value)>,
) -> Vec<(String, Value)> {
    entries
        .into_iter()
        .map(|(key, value)| (key.into(), value))
        .collect()
}

fn ordered_json(entries: &[(String, Value)]) -> String {
    let mut output = String::from("{");
    for (index, (key, value)) in entries.iter().enumerate() {
        if index > 0 {
            output.push(',');
        }
        output.push_str(&serde_json::to_string(key).unwrap());
        output.push(':');
        output.push_str(&js_json(value));
    }
    output.push('}');
    output
}

fn js_json(value: &Value) -> String {
    match value {
        Value::Null => "null".to_owned(),
        Value::Bool(value) => value.to_string(),
        Value::Number(value) => {
            let value = value.as_f64().unwrap_or_default();
            if value == 0.0 {
                "0".to_owned()
            } else {
                ryu_js::Buffer::new().format(value).to_owned()
            }
        }
        Value::String(value) => serde_json::to_string(value).unwrap(),
        Value::Array(values) => format!(
            "[{}]",
            values.iter().map(js_json).collect::<Vec<_>>().join(",")
        ),
        Value::Object(values) => format!(
            "{{{}}}",
            values
                .iter()
                .map(|(key, value)| format!(
                    "{}:{}",
                    serde_json::to_string(key).unwrap(),
                    js_json(value)
                ))
                .collect::<Vec<_>>()
                .join(",")
        ),
    }
}

impl OrderedValue {
    fn value(&self) -> Value {
        match self {
            Self::Null => Value::Null,
            Self::Bool(value) => Value::Bool(*value),
            Self::Number(value) => Value::Number(value.clone()),
            Self::String(value) => Value::String(value.clone()),
            Self::Array(values) => Value::Array(values.iter().map(Self::value).collect()),
            Self::Object(entries) => Value::Object(
                entries
                    .iter()
                    .map(|(key, value)| (key.clone(), value.value()))
                    .collect(),
            ),
        }
    }

    fn js_json(&self) -> String {
        match self {
            Self::Null => "null".to_owned(),
            Self::Bool(value) => value.to_string(),
            Self::Number(value) => {
                let value = value.as_f64().unwrap_or_default();
                if value == 0.0 {
                    "0".to_owned()
                } else {
                    ryu_js::Buffer::new().format(value).to_owned()
                }
            }
            Self::String(value) => serde_json::to_string(value).unwrap(),
            Self::Array(values) => format!(
                "[{}]",
                values
                    .iter()
                    .map(Self::js_json)
                    .collect::<Vec<_>>()
                    .join(",")
            ),
            Self::Object(entries) => format!(
                "{{{}}}",
                entries
                    .iter()
                    .map(|(key, value)| format!(
                        "{}:{}",
                        serde_json::to_string(key).unwrap(),
                        value.js_json()
                    ))
                    .collect::<Vec<_>>()
                    .join(",")
            ),
        }
    }

    fn insert_source_json(&self, output: &mut BTreeMap<String, String>) {
        output
            .entry(serde_json::to_string(&self.value()).unwrap())
            .or_insert_with(|| self.js_json());
    }

    fn collect_source_json(&self, output: &mut BTreeMap<String, String>) {
        match self {
            Self::Array(values) => {
                for value in values {
                    value.collect_source_json(output);
                }
            }
            Self::Object(entries) => {
                let node_type = entries.iter().find_map(|(key, value)| {
                    (key == "type")
                        .then_some(value)
                        .and_then(|value| match value {
                            Self::String(value) => Some(value.as_str()),
                            _ => None,
                        })
                });
                if matches!(
                    node_type,
                    Some("simpleField" | "complexField" | "shape" | "chart")
                ) {
                    self.insert_source_json(output);
                }
                if matches!(node_type, Some("inlineSdt" | "blockSdt"))
                    && let Some((_, properties)) =
                        entries.iter().find(|(key, _)| key == "properties")
                {
                    properties.insert_source_json(output);
                    if let Self::Object(properties) = properties {
                        for (_, value) in properties
                            .iter()
                            .filter(|(key, _)| matches!(key.as_str(), "listItems" | "dataBinding"))
                        {
                            value.insert_source_json(output);
                        }
                    }
                }
                for (_, value) in entries {
                    value.collect_source_json(output);
                }
            }
            _ => {}
        }
    }
}

fn needs_source_json(value: &Value) -> bool {
    match value {
        Value::Array(values) => values.iter().any(needs_source_json),
        Value::Object(values) => {
            matches!(
                string(values.get("type")),
                Some("simpleField" | "complexField" | "shape" | "chart" | "inlineSdt" | "blockSdt")
            ) || values.values().any(needs_source_json)
        }
        _ => false,
    }
}

fn source_json(value: &Value, values: &BTreeMap<String, String>) -> String {
    serde_json::to_string(value)
        .ok()
        .and_then(|key| values.get(&key).cloned())
        .unwrap_or_else(|| js_json(value))
}

fn drop_nulls(value: Value) -> Value {
    match value {
        Value::Array(values) => Value::Array(values.into_iter().map(drop_nulls).collect()),
        Value::Object(values) => Value::Object(
            values
                .into_iter()
                .filter(|(_, value)| !value.is_null())
                .map(|(key, value)| (key, drop_nulls(value)))
                .collect(),
        ),
        value => value,
    }
}

fn map_from_value(value: Value) -> JsonObject {
    match drop_nulls(value) {
        Value::Object(value) => value.into_iter().collect(),
        _ => JsonObject::new(),
    }
}

fn value_from_map(value: &JsonObject) -> Value {
    Value::Object(
        value
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect(),
    )
}

pub(crate) fn any_from_value(value: Value) -> Result<Any, String> {
    match value {
        Value::Null => Ok(Any::Null),
        Value::Bool(value) => Ok(Any::Bool(value)),
        Value::Number(value) if value.is_i64() => Ok(Any::from(value.as_i64().unwrap())),
        Value::Number(value) if value.is_u64() => Any::try_from(value.as_u64().unwrap())
            .map_err(|value| format!("JSON number {value} exceeds the yrs integer range")),
        Value::Number(value) => Ok(Any::Number(
            value
                .as_f64()
                .ok_or_else(|| format!("invalid JSON number {value}"))?,
        )),
        Value::String(value) => Ok(Any::String(Arc::from(value))),
        Value::Array(values) => values
            .into_iter()
            .map(any_from_value)
            .collect::<Result<Vec<_>, _>>()
            .map(Arc::from)
            .map(Any::Array),
        Value::Object(values) => {
            let mut entries = values
                .into_iter()
                .map(|(key, value)| Ok((key, any_from_value(value)?)))
                .collect::<Result<Vec<_>, String>>()?;
            entries.sort_unstable_by(|left, right| left.0.cmp(&right.0));
            Ok(Any::Map(Arc::new(entries.into_iter().collect())))
        }
    }
}

pub(crate) fn yrs_attrs(values: JsonObject) -> Result<Attrs, String> {
    let mut entries = values
        .into_iter()
        .map(|(key, value)| Ok((Arc::<str>::from(key), any_from_value(value)?)))
        .collect::<Result<Vec<_>, String>>()?;
    entries.sort_unstable_by(|left, right| left.0.cmp(&right.0));
    Ok(entries.into_iter().collect())
}

pub(crate) fn payload(values: JsonObject) -> Result<Vec<(String, Any)>, String> {
    values
        .into_iter()
        .map(|(key, value)| Ok((key, any_from_value(value)?)))
        .collect()
}

fn merge_plain(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    match (object(target), object(source)) {
        (None, None) => None,
        (Some(target), None) => Some(Value::Object(target.clone())),
        (None, Some(source)) => Some(Value::Object(source.clone())),
        (Some(target), Some(source)) => {
            let mut result = target.clone();
            for (key, value) in source {
                result.insert(key.clone(), value.clone());
            }
            Some(Value::Object(result))
        }
    }
}

fn merge_font_family(target: Option<&Value>, source: &Value) -> Value {
    let mut result = object(target).cloned().unwrap_or_default();
    let source = source.as_object().cloned().unwrap_or_default();
    for (explicit, theme) in [
        ("ascii", "asciiTheme"),
        ("hAnsi", "hAnsiTheme"),
        ("eastAsia", "eastAsiaTheme"),
        ("cs", "csTheme"),
    ] {
        if source.contains_key(explicit) || source.contains_key(theme) {
            result.remove(explicit);
            result.remove(theme);
            if let Some(value) = source.get(explicit) {
                result.insert(explicit.to_owned(), value.clone());
            }
            if let Some(value) = source.get(theme) {
                result.insert(theme.to_owned(), value.clone());
            }
        }
    }
    for (key, value) in source {
        if ![
            "ascii",
            "asciiTheme",
            "hAnsi",
            "hAnsiTheme",
            "eastAsia",
            "eastAsiaTheme",
            "cs",
            "csTheme",
        ]
        .contains(&key.as_str())
        {
            result.insert(key, value);
        }
    }
    Value::Object(result)
}

fn merge_text_formatting(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    let target_object = object(target);
    let source_object = object(source);
    if source_object.is_none() {
        return target.cloned();
    }
    if target_object.is_none() {
        return source.cloned();
    }
    let target_object = target_object.unwrap();
    let mut result = target_object.clone();
    for (key, value) in source_object.unwrap() {
        if key == "fontFamily" && value.is_object() {
            result.insert(
                key.clone(),
                merge_font_family(target_object.get(key), value),
            );
        } else if key == "color" && value.is_object() {
            let explicit = truthy(field(Some(value), "rgb"))
                || truthy(field(Some(value), "themeColor"))
                || truthy(field(Some(value), "themeTint"))
                || truthy(field(Some(value), "themeShade"));
            if !truthy(field(Some(value), "auto")) || explicit {
                result.insert(key.clone(), value.clone());
            }
        } else if value.is_object() {
            result.insert(
                key.clone(),
                merge_plain(target_object.get(key), Some(value)).unwrap(),
            );
        } else {
            result.insert(key.clone(), value.clone());
        }
    }
    Some(Value::Object(result))
}

fn merge_paragraph_formatting(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    let Some(source) = object(source) else {
        return target.cloned();
    };
    let mut result = object(target).cloned().unwrap_or_default();
    if let Some(value) = source
        .get("indentFirstLine")
        .filter(|value| !value.is_null())
    {
        result.insert("indentFirstLine".to_owned(), value.clone());
        match source.get("hangingIndent").filter(|value| !value.is_null()) {
            Some(hanging) => {
                result.insert("hangingIndent".to_owned(), hanging.clone());
            }
            None => {
                result.remove("hangingIndent");
            }
        }
    }
    for (key, value) in source {
        if key == "runProperties" {
            if let Some(merged) = merge_text_formatting(result.get(key), Some(value)) {
                result.insert(key.clone(), merged);
            }
        } else if ["borders", "numPr", "frame"].contains(&key.as_str()) {
            result.insert(
                key.clone(),
                merge_plain(result.get(key), Some(value)).unwrap_or_else(|| value.clone()),
            );
        } else if matches!(key.as_str(), "indentFirstLine" | "hangingIndent") {
            continue;
        } else {
            result.insert(key.clone(), value.clone());
        }
    }
    Some(Value::Object(result))
}

impl StyleResolver {
    fn new(definitions: Option<&Value>) -> Self {
        let Some(definitions) = object(definitions) else {
            return Self::default();
        };
        let mut resolver = Self {
            enabled: true,
            doc_defaults: definitions.get("docDefaults").cloned(),
            ..Self::default()
        };
        for style in array(definitions.get("styles")) {
            let Some(style_id) = string(field(Some(style), "styleId")) else {
                continue;
            };
            resolver.styles.insert(style_id.to_owned(), style.clone());
        }
        resolver.default_paragraph = resolver.find_default("paragraph").or_else(|| {
            resolver
                .styles
                .contains_key("Normal")
                .then(|| "Normal".to_owned())
        });
        resolver.default_table = resolver.find_default("table");
        resolver.default_character = resolver.find_default("character");
        resolver
    }

    fn find_default(&self, style_type: &str) -> Option<String> {
        self.styles.iter().find_map(|(id, style)| {
            (string(field(Some(style), "type")) == Some(style_type)
                && truthy(field(Some(style), "default")))
            .then(|| id.clone())
        })
    }

    fn style(&self, style_id: &str) -> Option<&Value> {
        self.styles.get(style_id)
    }

    fn default_style(&self, style_type: &str) -> Option<&Value> {
        let id = match style_type {
            "paragraph" => self.default_paragraph.as_deref(),
            "table" => self.default_table.as_deref(),
            "character" => self.default_character.as_deref(),
            _ => None,
        };
        id.and_then(|id| self.style(id))
    }

    fn resolve_paragraph_style(&self, style_id: Option<&str>) -> (Option<Value>, Option<Value>) {
        let mut paragraph = merge_paragraph_formatting(
            field(self.doc_defaults.as_ref(), "pPr"),
            self.table_paragraph_formatting.as_ref(),
        );
        let mut run = field(self.doc_defaults.as_ref(), "rPr").cloned();
        let style = style_id
            .and_then(|id| self.style(id))
            .or_else(|| self.default_style("paragraph"));
        if let Some(style) = style {
            paragraph = merge_paragraph_formatting(paragraph.as_ref(), field(Some(style), "pPr"));
            run = merge_text_formatting(run.as_ref(), field(Some(style), "rPr"));
        }
        if style_id.is_some() && style.is_none() {
            if let Some(style) = self.default_style("paragraph") {
                paragraph =
                    merge_paragraph_formatting(paragraph.as_ref(), field(Some(style), "pPr"));
                run = merge_text_formatting(run.as_ref(), field(Some(style), "rPr"));
            }
        }
        if style_id.is_none() && style.is_none() && self.doc_defaults.is_none() {
            paragraph = merge_paragraph_formatting(
                Some(&json!({
                    "spaceAfter": 160,
                    "lineSpacing": 259,
                    "lineSpacingRule": "auto"
                })),
                self.table_paragraph_formatting.as_ref(),
            );
        }
        (paragraph, run)
    }

    fn resolve_run_style(&self, style_id: Option<&str>) -> Option<Value> {
        let mut result = field(self.doc_defaults.as_ref(), "rPr").cloned();
        result = merge_text_formatting(
            result.as_ref(),
            self.default_style("character")
                .and_then(|style| field(Some(style), "rPr")),
        );
        if let Some(style) = style_id.and_then(|id| self.style(id)) {
            result = merge_text_formatting(result.as_ref(), field(Some(style), "rPr"));
        }
        result
    }

    fn run_style_own(&self, style_id: Option<&str>) -> Option<Value> {
        style_id
            .and_then(|id| self.style(id))
            .and_then(|style| field(Some(style), "rPr"))
            .cloned()
    }
}

fn mark(name: &str, attrs: Vec<(String, Value)>) -> Mark {
    Mark {
        name: name.to_owned(),
        attrs,
    }
}

fn formatting_to_marks(formatting: Option<&Value>) -> Vec<Mark> {
    let mut marks = Vec::new();
    let Some(formatting) = object(formatting) else {
        return marks;
    };
    if truthy(formatting.get("bold")) {
        marks.push(mark("bold", vec![]));
    }
    if truthy(formatting.get("italic")) {
        marks.push(mark("italic", vec![]));
    }
    if let Some(underline) = object(formatting.get("underline"))
        && string(underline.get("style")) != Some("none")
    {
        marks.push(mark(
            "underline",
            ordered_object([
                (
                    "style",
                    underline.get("style").cloned().unwrap_or(Value::Null),
                ),
                ("color", nullish(underline.get("color"))),
            ]),
        ));
    }
    if truthy(formatting.get("strike")) || truthy(formatting.get("doubleStrike")) {
        marks.push(mark(
            "strike",
            ordered_object([(
                "double",
                Value::Bool(truthy(formatting.get("doubleStrike"))),
            )]),
        ));
    }
    if let Some(color) = object(formatting.get("color"))
        && !truthy(color.get("auto"))
    {
        marks.push(mark(
            "textColor",
            ordered_object([
                ("rgb", nullish(color.get("rgb"))),
                ("themeColor", nullish(color.get("themeColor"))),
                ("themeTint", nullish(color.get("themeTint"))),
                ("themeShade", nullish(color.get("themeShade"))),
            ]),
        ));
    }
    let shading_fill =
        object(formatting.get("shading")).and_then(|shading| object(shading.get("fill")));
    let shading_highlight = shading_fill.and_then(|fill| {
        let pattern =
            string(object(formatting.get("shading")).and_then(|shading| shading.get("pattern")));
        (pattern.is_none() || pattern == Some("clear"))
            .then(|| string(fill.get("rgb")))
            .flatten()
            .filter(|_| !truthy(fill.get("auto")))
            .map(|rgb| {
                if rgb.starts_with('#') {
                    rgb.to_owned()
                } else {
                    format!("#{rgb}")
                }
            })
    });
    let highlight = string(formatting.get("highlight"))
        .filter(|value| *value != "none")
        .map(str::to_owned)
        .or(shading_highlight);
    if let Some(highlight) = highlight {
        marks.push(mark(
            "highlight",
            ordered_object([("color", Value::String(highlight))]),
        ));
    }
    if formatting.contains_key("fontSize") || formatting.contains_key("fontSizeCs") {
        marks.push(mark(
            "fontSize",
            ordered_object([
                ("size", nullish(formatting.get("fontSize"))),
                ("sizeCs", nullish(formatting.get("fontSizeCs"))),
            ]),
        ));
    }
    if let Some(font) = object(formatting.get("fontFamily")) {
        marks.push(mark(
            "fontFamily",
            ordered_object([
                ("ascii", nullish(font.get("ascii"))),
                ("hAnsi", nullish(font.get("hAnsi"))),
                ("eastAsia", nullish(font.get("eastAsia"))),
                ("cs", nullish(font.get("cs"))),
                ("asciiTheme", nullish(font.get("asciiTheme"))),
                ("hAnsiTheme", nullish(font.get("hAnsiTheme"))),
                ("eastAsiaTheme", nullish(font.get("eastAsiaTheme"))),
                ("csTheme", nullish(font.get("csTheme"))),
            ]),
        ));
    }
    if let Some(language) = object(formatting.get("language")) {
        marks.push(mark(
            "language",
            ordered_object([
                ("latin", nullish(language.get("latin"))),
                ("eastAsia", nullish(language.get("eastAsia"))),
                ("bidi", nullish(language.get("bidi"))),
            ]),
        ));
    }
    match string(formatting.get("vertAlign")) {
        Some("superscript") => marks.push(mark("superscript", vec![])),
        Some("subscript") => marks.push(mark("subscript", vec![])),
        _ => {}
    }
    for (key, name) in [
        ("allCaps", "allCaps"),
        ("smallCaps", "smallCaps"),
        ("emboss", "emboss"),
        ("imprint", "imprint"),
        ("shadow", "textShadow"),
        ("outline", "textOutline"),
        ("hidden", "hidden"),
        ("rtl", "rtl"),
    ] {
        if truthy(formatting.get(key)) {
            marks.push(mark(name, vec![]));
        }
    }
    // Document-grid opt-out (w:snapToGrid, default on): only an authored off
    // becomes a mark, mirroring documentToYrs.
    if formatting.get("snapToGrid") == Some(&Value::Bool(false)) {
        marks.push(mark("snapToGrid", vec![]));
    }
    if ["spacing", "position", "scale", "kerning"]
        .iter()
        .any(|key| formatting.contains_key(*key))
    {
        marks.push(mark(
            "characterSpacing",
            ordered_object([
                ("spacing", nullish(formatting.get("spacing"))),
                ("position", nullish(formatting.get("position"))),
                ("scale", nullish(formatting.get("scale"))),
                ("kerning", nullish(formatting.get("kerning"))),
            ]),
        ));
    }
    if let Some(value) = string(formatting.get("emphasisMark")).filter(|value| *value != "none") {
        marks.push(mark(
            "emphasisMark",
            ordered_object([("type", Value::String(value.to_owned()))]),
        ));
    }
    if let Some(value) = string(formatting.get("effect")).filter(|value| *value != "none") {
        marks.push(mark(
            "textEffect",
            ordered_object([("effect", Value::String(value.to_owned()))]),
        ));
    }
    if let Some(value) = formatting.get("modernEffects") {
        marks.push(mark(
            "modernTextEffects",
            ordered_object([("effects", value.clone())]),
        ));
    }
    if let Some(value) = formatting.get("styleId") {
        marks.push(mark(
            "runStyle",
            ordered_object([("styleId", value.clone())]),
        ));
    }
    marks
}

fn mark_attrs(mark: &Mark) -> Value {
    Value::Object(mark.attrs.iter().cloned().collect())
}

fn marks_to_attrs(marks: &[Mark]) -> JsonObject {
    let boolean_marks = [
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
    let mut attrs = JsonObject::new();
    for mark in marks {
        if mark.name == "comment" || mark.name == "footnoteRef" {
            continue;
        }
        if boolean_marks.contains(&mark.name.as_str()) {
            attrs.insert(mark.name.clone(), Value::Bool(true));
        } else if mark.name == "snapToGrid" {
            attrs.insert("snapToGrid".to_owned(), Value::Bool(false));
        } else if mark.name == "highlight" {
            attrs.insert(
                "highlight".to_owned(),
                mark.attrs
                    .iter()
                    .find(|(key, _)| key == "color")
                    .map(|(_, value)| value.clone())
                    .unwrap_or(Value::Null),
            );
        } else if mark.name == "insertion" || mark.name == "deletion" {
            let get = |name: &str| {
                mark.attrs
                    .iter()
                    .find(|(key, _)| key == name)
                    .map(|(_, value)| value.clone())
                    .unwrap_or(Value::Null)
            };
            attrs.insert(
                if mark.name == "insertion" {
                    "ins".to_owned()
                } else {
                    "del".to_owned()
                },
                drop_nulls(json!({
                    "id": get("revisionId"),
                    "author": get("author"),
                    "date": get("date")
                })),
            );
        } else {
            attrs.insert(mark.name.clone(), drop_nulls(mark_attrs(mark)));
        }
    }
    attrs
}

fn marks_key(marks: &[Mark]) -> String {
    let mut values: Vec<String> = marks
        .iter()
        .filter(|mark| mark.name != "hyperlink" && mark.name != "comment")
        .map(|mark| format!("{}:{}", mark.name, ordered_json(&mark.attrs)))
        .collect();
    values.sort();
    values.join("|")
}

fn with_mark(marks: &[Mark], next: Mark) -> Vec<Mark> {
    let name = next.name.clone();
    marks
        .iter()
        .filter(|mark| mark.name != name)
        .cloned()
        .chain(std::iter::once(next))
        .collect()
}

fn text_unit(text: String, marks: &[Mark]) -> InlineUnit {
    InlineUnit {
        pm_size: utf16_len(&text),
        content: UnitContent::Text(text),
        attrs: marks_to_attrs(marks),
        comments: Vec::new(),
        marks: marks.to_vec(),
    }
}

fn embed_unit(kind: &str, payload: JsonObject, marks: &[Mark], pm_size: u32) -> InlineUnit {
    InlineUnit {
        content: UnitContent::Embed {
            kind: kind.to_owned(),
            payload,
        },
        attrs: marks_to_attrs(marks),
        pm_size,
        comments: Vec::new(),
        marks: marks.to_vec(),
    }
}

fn run_format_marks(
    run: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
) -> Vec<Mark> {
    let formatting = field(Some(run), "formatting");
    let style_id = string(field(formatting, "styleId"));
    let run_style = styles.run_style_own(style_id);
    let inherited = merge_text_formatting(style_formatting, run_style.as_ref());
    let merged = merge_text_formatting(inherited.as_ref(), formatting);
    let mut marks = formatting_to_marks(merged.as_ref());
    let hyperlink_style = style_id.is_some_and(is_hyperlink_style_name)
        || style_id
            .and_then(|id| styles.style(id))
            .and_then(|style| string(field(Some(style), "name")))
            .is_some_and(is_hyperlink_style_name);
    if hyperlink_style {
        for (property, name) in [("color", "textColor"), ("underline", "underline")] {
            if field(formatting, property).is_none()
                && field(run_style.as_ref(), property).is_some()
                && let Some(mark) = marks.iter_mut().find(|mark| mark.name == name)
            {
                mark.attrs
                    .push(("inheritedHyperlink".to_owned(), Value::Bool(true)));
            }
        }
    }
    marks
}

fn run_marks(run: &Value, style_formatting: Option<&Value>, styles: &StyleResolver) -> Vec<Mark> {
    let mut marks = run_format_marks(run, style_formatting, styles);
    let changes = array(field(Some(run), "propertyChanges"));
    if !changes.is_empty() {
        let records = changes
            .iter()
            .map(|change| {
                let mut record = map_from_value(change.clone());
                for (source, target) in [
                    ("previousFormatting", "previousAttributes"),
                    ("currentFormatting", "currentAttributes"),
                ] {
                    let before = json!({"formatting": change[source]});
                    record.insert(
                        target.into(),
                        value_from_map(&marks_to_attrs(&run_format_marks(
                            &before,
                            style_formatting,
                            styles,
                        ))),
                    );
                }
                value_from_map(&record)
            })
            .collect();
        marks.push(mark(
            "rPrChange",
            vec![("changes".into(), Value::Array(records))],
        ));
    }
    marks
}

fn is_hyperlink_style_name(name: &str) -> bool {
    name.eq_ignore_ascii_case("Hyperlink") || name.eq_ignore_ascii_case("FollowedHyperlink")
}

fn emu_to_pixels(value: f64) -> f64 {
    value / 914_400.0 * 96.0
}

fn image_payload(image: &Value) -> JsonObject {
    let size = field(Some(image), "size");
    let wrap = field(Some(image), "wrap");
    let position = field(Some(image), "position");
    let transform = field(Some(image), "transform");
    let outline = field(Some(image), "outline");
    let wrap_type = string(field(wrap, "type")).unwrap_or_default();
    let wrap_text = string(field(wrap, "wrapText"));
    let horizontal = field(position, "horizontal");
    let vertical = field(position, "vertical");
    let alignment = string(field(horizontal, "alignment"));
    let css_float = if wrap_type == "inline" || wrap_type == "topAndBottom" {
        "none"
    } else if ["square", "tight", "through"].contains(&wrap_type) {
        if wrap_text == Some("left") {
            "right"
        } else if wrap_text == Some("right") {
            "left"
        } else if matches!(alignment, Some("left" | "right")) {
            alignment.unwrap()
        } else {
            "none"
        }
    } else {
        "none"
    };
    let display_mode = if wrap_type == "inline" {
        "inline"
    } else if wrap_type == "topAndBottom" {
        "block"
    } else if matches!(wrap_type, "behind" | "inFront") || css_float != "none" {
        "float"
    } else {
        "block"
    };
    let mut transforms = Vec::new();
    if let Some(rotation) = number(field(transform, "rotation")).filter(|value| *value != 0.0) {
        transforms.push(format!("rotate({rotation}deg)"));
    }
    if truthy(field(transform, "flipH")) {
        transforms.push("scaleX(-1)".to_owned());
    }
    if truthy(field(transform, "flipV")) {
        transforms.push("scaleY(-1)".to_owned());
    }
    let outline_width = number(field(outline, "width")).filter(|value| *value != 0.0);
    let border_width =
        outline_width.map(|value| (value / 914_400.0 * 96.0 * 100.0).round() / 100.0);
    let border_color =
        string(field(field(outline, "color"), "rgb")).map(|value| format!("#{value}"));
    let border_style = outline_width.map(|_| match string(field(outline, "style")) {
        Some("dot" | "sysDot") => "dotted",
        Some(
            "dash" | "lgDash" | "dashDot" | "lgDashDot" | "lgDashDotDot" | "sysDash" | "sysDashDot"
            | "sysDashDotDot",
        ) => "dashed",
        _ => "solid",
    });
    let axis = |axis: Option<&Value>| {
        axis.map(|axis| {
            json!({
                "relativeTo": nullish(field(Some(axis), "relativeTo")),
                "posOffset": nullish(field(Some(axis), "posOffset")),
                "align": nullish(field(Some(axis), "alignment"))
            })
        })
    };
    map_from_value(json!({
        "src": string(field(Some(image), "src")).unwrap_or_default(),
        "alt": nullish(field(Some(image), "alt")),
        "title": nullish(field(Some(image), "title")),
        "width": number(field(size, "width")).filter(|value| *value != 0.0).map(emu_to_pixels),
        "height": number(field(size, "height")).filter(|value| *value != 0.0).map(emu_to_pixels),
        "rId": nullish(field(Some(image), "rId")),
        "wrapType": wrap_type,
        "displayMode": display_mode,
        "cssFloat": css_float,
        "transform": (!transforms.is_empty()).then(|| transforms.join(" ")),
        "distTop": number(field(wrap, "distT")).map(emu_to_pixels),
        "distBottom": number(field(wrap, "distB")).map(emu_to_pixels),
        "distLeft": number(field(wrap, "distL")).map(emu_to_pixels),
        "distRight": number(field(wrap, "distR")).map(emu_to_pixels),
        "position": position.map(|_| json!({
            "horizontal": axis(horizontal),
            "vertical": axis(vertical),
            "relativeHeight": nullish(field(position, "relativeHeight"))
        })),
        "borderWidth": border_width,
        "borderColor": border_color,
        "borderStyle": border_style,
        "wrapText": wrap_text,
        "hlinkHref": nullish(field(Some(image), "hlinkHref")),
        "cropTop": nullish(field(field(Some(image), "crop"), "top")),
        "cropRight": nullish(field(field(Some(image), "crop"), "right")),
        "cropBottom": nullish(field(field(Some(image), "crop"), "bottom")),
        "cropLeft": nullish(field(field(Some(image), "crop"), "left")),
        "shapeType": nullish(field(Some(image), "shapeType")),
        "opacity": nullish(field(Some(image), "opacity")),
        "effectExtentTop": number(field(field(Some(image), "padding"), "top"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "effectExtentBottom": number(field(field(Some(image), "padding"), "bottom"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "effectExtentLeft": number(field(field(Some(image), "padding"), "left"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "effectExtentRight": number(field(field(Some(image), "padding"), "right"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "layoutInCell": nullish(field(Some(image), "layoutInCell")),
        "allowOverlap": nullish(field(Some(image), "allowOverlap"))
    }))
}

/// Carries the authored XML a picture or shape replays on save while unedited.
fn with_source_xml(mut payload: JsonObject, content: &Value) -> JsonObject {
    if let Some(xml) = field(Some(content), "sourceXml").filter(|xml| xml.is_string()) {
        payload.insert("sourceXml".to_owned(), xml.clone());
    }
    payload
}

fn shape_payload(shape: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    map_from_value(json!({ "shapeJson": source_json(shape, source) }))
}

fn chart_payload(chart: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    let size = field(Some(chart), "size");
    map_from_value(json!({
        "chartJson": source_json(chart, source),
        "chartType": nullish(field(Some(chart), "chartType")),
        "title": nullish(field(Some(chart), "title")),
        "width": number(field(size, "width")).filter(|value| *value != 0.0).map(emu_to_pixels).unwrap_or(320.0),
        "height": number(field(size, "height")).filter(|value| *value != 0.0).map(emu_to_pixels).unwrap_or(220.0),
        "rId": nullish(field(Some(chart), "rId")),
        "path": nullish(field(Some(chart), "path"))
    }))
}

pub(crate) fn numeric_field_instruction(instruction: &str) -> bool {
    let instruction = instruction.trim();
    !instruction.is_empty() && instruction.bytes().all(|byte| byte.is_ascii_digit())
}

fn field_payload(
    field_value: &Value,
    style_formatting: Option<&Value>,
    source: &BTreeMap<String, String>,
) -> (JsonObject, Vec<Mark>) {
    let kind = string(field(Some(field_value), "type")).unwrap_or_default();
    let runs = if kind == "simpleField" {
        array(field(Some(field_value), "content"))
    } else {
        array(field(Some(field_value), "fieldResult"))
    };
    let mut display_text = String::new();
    let mut field_formatting = None;
    for child in runs {
        if string(field(Some(child), "type")) != Some("run") {
            continue;
        }
        for content in array(field(Some(child), "content")) {
            if string(field(Some(content), "type")) == Some("text") {
                display_text.push_str(string(field(Some(content), "text")).unwrap_or_default());
            }
        }
        if field_formatting.is_none() {
            field_formatting = field(Some(child), "formatting");
        }
    }
    let formatting = field_formatting.or_else(|| {
        (kind == "complexField")
            .then(|| field(Some(field_value), "formatting"))
            .flatten()
    });
    let merged = merge_text_formatting(style_formatting, formatting);
    let mut payload = map_from_value(json!({
        "fieldType": nullish(field(Some(field_value), "fieldType")),
        "instruction": nullish(field(Some(field_value), "instruction")),
        "displayText": display_text,
        "fieldKind": if kind == "simpleField" { "simple" } else { "complex" },
        "fldLock": boolean(field(Some(field_value), "fldLock")).unwrap_or(false),
        "dirty": boolean(field(Some(field_value), "dirty")).unwrap_or(false),
        "displayMode": string(field(field(Some(field_value), "fieldTree"), "displayMode")).unwrap_or("result"),
        "hasCachedResult": !display_text.is_empty(),
        "fieldData": source_json(field_value, source),
        "modelKind": "field"
    }));
    if let Some(id) = field(field(Some(field_value), "continuation"), "id") {
        payload.insert("continuationId".into(), id.clone());
    }
    (payload, formatting_to_marks(merged.as_ref()))
}

fn math_payload(math: &Value) -> JsonObject {
    map_from_value(json!({
        "display": nullish(field(Some(math), "display")),
        "ommlXml": nullish(field(Some(math), "ommlXml")),
        "plainText": string(field(Some(math), "plainText")).unwrap_or_default()
    }))
}

fn hyperlink_mark(hyperlink: &Value) -> Mark {
    let href = string(field(Some(hyperlink), "href"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
        .or_else(|| {
            string(field(Some(hyperlink), "anchor"))
                .filter(|value| !value.is_empty())
                .map(|value| format!("#{value}"))
        })
        .unwrap_or_default();
    mark(
        "hyperlink",
        ordered_object([
            ("href", Value::String(href)),
            ("tooltip", nullish(field(Some(hyperlink), "tooltip"))),
            ("rId", nullish(field(Some(hyperlink), "rId"))),
        ]),
    )
}

fn note_ref_unit(id: &Value, note_type: &str, marks: &[Mark]) -> InlineUnit {
    let note_mark = mark(
        "footnoteRef",
        ordered_object([
            ("id", Value::String(js_string(id))),
            ("noteType", Value::String(note_type.to_owned())),
        ]),
    );
    let all_marks: Vec<Mark> = marks
        .iter()
        .cloned()
        .chain(std::iter::once(note_mark))
        .collect();
    embed_unit(
        "noteRef",
        map_from_value(if note_type == "endnote" {
            json!({ "endnoteRefId": id })
        } else {
            json!({ "footnoteRefId": id })
        }),
        &all_marks,
        1,
    )
}

fn hidden_marks(marks: &[Mark]) -> &[Mark] {
    marks
        .iter()
        .find(|mark| mark.name == "hidden")
        .map_or(&[], std::slice::from_ref)
}

fn run_content_to_units(
    content: &Value,
    marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    match string(field(Some(content), "type")).unwrap_or_default() {
        "text" => string(field(Some(content), "text"))
            .filter(|text| !text.is_empty())
            .map(|text| vec![text_unit(text.to_owned(), marks)])
            .unwrap_or_default(),
        "tab" => vec![text_unit("\t".to_owned(), marks)],
        "break" if flow_break_type(content).is_some() => vec![embed_unit(
            if flow_break_type(content) == Some("column") {
                "columnBreak"
            } else {
                "pageBreak"
            },
            JsonObject::new(),
            marks,
            1,
        )],
        "break"
            if string(field(Some(content), "breakType"))
                .is_none_or(|kind| kind == "textWrapping") =>
        {
            vec![embed_unit("break", JsonObject::new(), marks, 1)]
        }
        "softHyphen" => vec![text_unit("\u{00ad}".to_owned(), marks)],
        "noBreakHyphen" => vec![text_unit("\u{2011}".to_owned(), marks)],
        "symbol" => {
            let Some(codepoint) = string(field(Some(content), "char"))
                .and_then(|value| u32::from_str_radix(value, 16).ok())
                .and_then(char::from_u32)
            else {
                return vec![];
            };
            let font = string(field(Some(content), "font"))
                .filter(|value| !value.is_empty())
                .map(|value| Value::String(value.to_owned()))
                .unwrap_or(Value::Null);
            let symbol_mark = mark(
                "fontFamily",
                ordered_object([
                    ("ascii", font.clone()),
                    ("hAnsi", font.clone()),
                    ("eastAsia", font.clone()),
                    ("cs", font),
                    ("asciiTheme", Value::Null),
                    ("hAnsiTheme", Value::Null),
                    ("eastAsiaTheme", Value::Null),
                    ("csTheme", Value::Null),
                ]),
            );
            vec![text_unit(
                codepoint.to_string(),
                &with_mark(marks, symbol_mark),
            )]
        }
        "commentReference" => {
            let mut value = json!({
                "fieldType": "COMMENT",
                "instruction": "",
                "displayText": "",
                "fieldKind": "simple",
                "fldLock": false,
                "dirty": false,
                "displayMode": "result",
                "hasCachedResult": false,
                "modelKind": "commentReference"
            });
            if let Some(id) = field(Some(content), "id") {
                value
                    .as_object_mut()
                    .unwrap()
                    .insert("commentId".to_owned(), id.clone());
            }
            vec![embed_unit("field", map_from_value(value), &[], 1)]
        }
        "drawing" => vec![embed_unit(
            "image",
            with_source_xml(
                image_payload(field(Some(content), "image").unwrap_or(&Value::Null)),
                content,
            ),
            hidden_marks(marks),
            1,
        )],
        "horizontalRule" => vec![embed_unit(
            "horizontalRule",
            map_from_value(json!({"rule": field(Some(content), "rule")})),
            marks,
            1,
        )],
        "shape" => vec![embed_unit(
            "shape",
            with_source_xml(
                shape_payload(
                    field(Some(content), "shape").unwrap_or(&Value::Null),
                    source,
                ),
                content,
            ),
            hidden_marks(marks),
            1,
        )],
        "chart" => vec![embed_unit(
            "chart",
            chart_payload(
                field(Some(content), "chart").unwrap_or(&Value::Null),
                source,
            ),
            hidden_marks(marks),
            1,
        )],
        "opaqueDrawing" => vec![embed_unit(
            "opaqueDrawing",
            map_from_value(json!({
                "kind": field(Some(content), "kind"),
                "xml": field(Some(content), "xml"),
            })),
            hidden_marks(marks),
            1,
        )],
        "footnoteRef" => field(Some(content), "id")
            .map(|id| note_ref_unit(id, "footnote", marks))
            .into_iter()
            .collect(),
        "endnoteRef" => field(Some(content), "id")
            .map(|id| note_ref_unit(id, "endnote", marks))
            .into_iter()
            .collect(),
        _ => vec![],
    }
}

fn run_to_units(
    run: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    extra_marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let marks: Vec<Mark> = run_marks(run, style_formatting, styles)
        .into_iter()
        .chain(extra_marks.iter().cloned())
        .collect();
    array(field(Some(run), "content"))
        .iter()
        .flat_map(|content| {
            if string(field(Some(content), "type")) == Some("fieldChar")
                && let Some(id) = field(Some(content), "continuationId")
            {
                return vec![embed_unit("bookmark", map_from_value(json!({
                    "id": id, "kind": format!("field{}", string(field(Some(content), "charType")).unwrap_or_default()),
                    "run": {"type": "run", "content": [content], "formatting": field(Some(run), "formatting")}
                })), &[], 0)];
            }
            run_content_to_units(content, &marks, source)
        })
        .collect()
}

fn hyperlink_to_units(
    hyperlink: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    extra_marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let marks: Vec<_> = extra_marks
        .iter()
        .cloned()
        .chain([hyperlink_mark(hyperlink)])
        .collect();
    let children =
        field(Some(hyperlink), "structuredChildren").or_else(|| field(Some(hyperlink), "children"));
    array(children)
        .iter()
        .flat_map(|child| inline_to_units(child, style_formatting, styles, &marks, source))
        .collect()
}

/// The same recursive inline grammar inside links, controls and revisions.
fn inline_to_units(
    child: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    extra_marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let mut units = match string(field(Some(child), "type")) {
        Some("run") => run_to_units(child, style_formatting, styles, &[], source),
        Some("hyperlink") => hyperlink_to_units(child, style_formatting, styles, &[], source),
        Some("simpleField" | "complexField") => {
            let (payload, marks) = field_payload(child, style_formatting, source);
            vec![embed_unit("field", payload, &marks, 1)]
        }
        Some("inlineSdt") => vec![embed_unit(
            "sdt",
            sdt_payload(child, style_formatting, styles, source),
            &[],
            1,
        )],
        Some("mathEquation") => vec![embed_unit("math", math_payload(child), &[], 1)],
        Some(kind @ ("bookmarkStart" | "bookmarkEnd")) => {
            let mut data = map_from_value(child.clone());
            data.remove("type");
            data.remove("position");
            data.insert(
                "kind".into(),
                json!(if kind == "bookmarkStart" {
                    "start"
                } else {
                    "end"
                }),
            );
            vec![embed_unit("bookmark", data, &[], 0)]
        }
        Some("insertion" | "deletion" | "moveFrom" | "moveTo") => {
            tracked_to_units(child, style_formatting, styles, source)
        }
        _ => Vec::new(),
    };
    for unit in &mut units {
        for mark in extra_marks {
            unit.marks = with_mark(&unit.marks, mark.clone());
        }
        unit.attrs = marks_to_attrs(&unit.marks);
        if let UnitContent::Embed { kind, payload } = &mut unit.content
            && kind == "bookmark"
            && unit.attrs.contains_key("hyperlink")
        {
            payload.insert("inHyperlink".into(), Value::Bool(true));
        }
    }
    units
}

fn field_to_units(
    value: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
    projection_id: usize,
) -> Vec<InlineUnit> {
    let result = array(field(field(Some(value), "structuredResult"), "inline"));
    let code = array(field(field(Some(value), "structuredCode"), "inline"));
    let projected_children: Vec<_> = code
        .iter()
        .enumerate()
        .map(|(index, child)| (-(index as isize) - 1, child))
        .chain(
            result
                .iter()
                .enumerate()
                .map(|(index, child)| (index as isize, child)),
        )
        .collect();
    if numeric_field_instruction(string(field(Some(value), "instruction")).unwrap_or_default())
        || string(field(Some(value), "type")) != Some("complexField")
        || !projected_children.iter().any(|(_, child)| {
            matches!(
                string(field(Some(child), "type")),
                Some("hyperlink" | "simpleField")
            )
        })
    {
        let (payload, marks) = field_payload(value, style_formatting, source);
        return vec![embed_unit("field", payload, &marks, 1)];
    }
    let mut units = Vec::new();
    let mut children = Vec::new();
    for (index, child) in projected_children {
        let mut projected = match string(field(Some(child), "type")) {
            Some("hyperlink") => hyperlink_to_units(child, style_formatting, styles, &[], source),
            Some("simpleField") => {
                let (payload, marks) = field_payload(child, style_formatting, source);
                vec![embed_unit("field", payload, &marks, 1)]
            }
            _ => continue,
        };
        let items: Vec<Value> = projected.iter().map(|unit| match &unit.content {
            UnitContent::Text(text) => json!({"kind":"text", "text":text, "attributes":unit.attrs}),
            UnitContent::Embed {kind, payload} => json!({"kind":"embed", "embedKind":kind, "payload":payload, "attributes":unit.attrs}),
        }).collect();
        children.push(json!({"index":index, "items":items}));
        for unit in &mut projected {
            unit.attrs.insert(
                "fieldResult".to_owned(),
                json!({"id":projection_id, "index":index}),
            );
        }
        units.extend(projected);
    }
    // The field shows the result its projected children leave: runs, and a
    // kept change's or content control's text.
    let mut visible = value.clone();
    visible["fieldResult"] = Value::Array(
        result
            .iter()
            .filter(|child| {
                matches!(
                    string(field(Some(child), "type")),
                    Some("run" | "rawXml" | "inlineSdt")
                )
            })
            .flat_map(|child| shown_runs(std::slice::from_ref(child)))
            .collect(),
    );
    let (mut payload, marks) = field_payload(&visible, style_formatting, source);
    payload.insert(
        "fieldData".to_owned(),
        Value::String(source_json(value, source)),
    );
    payload.insert(
        "resultProjection".to_owned(),
        json!({"id":projection_id, "children":children}),
    );
    units.push(embed_unit("field", payload, &marks, 1));
    units
}

/// A field's story units as the seed makes them in a paragraph styled
/// `style_id` in `package`: its projected children, each
/// marked with `fieldResult`, then its embed. An inline content control's
/// field (`project` false) is its embed alone, as `sdt_payload` seeds it.
/// Each unit is its text or embed kind and payload, and its attributes.
pub(crate) fn field_units(
    value: &Value,
    package: Option<&PackageContext>,
    style_id: Option<&str>,
    projection_id: usize,
    project: bool,
) -> Vec<(Result<String, (String, JsonObject)>, JsonObject)> {
    let none = StyleResolver::new(None);
    let styles = package.map_or(&none, |package| &package.styles);
    let paragraph = json!({ "formatting": { "styleId": style_id } });
    let style_formatting = paragraph_style_formatting(&paragraph, styles, None);
    let source = BTreeMap::new();
    let units = if project {
        field_to_units(
            value,
            style_formatting.as_ref(),
            styles,
            &source,
            projection_id,
        )
    } else {
        let (payload, marks) = field_payload(value, style_formatting.as_ref(), &source);
        vec![embed_unit("field", payload, &marks, 1)]
    };
    units
        .into_iter()
        .map(|unit| {
            let content = match unit.content {
                UnitContent::Text(text) => Ok(text),
                UnitContent::Embed { kind, payload } => Err((kind, payload)),
            };
            (content, unit.attrs)
        })
        .collect()
}

/// The runs field result `nodes` show (`shown_runs` in docx-parse's paragraph module).
pub(crate) fn shown_runs(nodes: &[Value]) -> Vec<Value> {
    nodes
        .iter()
        .flat_map(|node| match string(field(Some(node), "type")) {
            Some("run") => vec![node.clone()],
            Some("hyperlink") => shown_runs(array(field(Some(node), "children"))),
            Some("simpleField") => array(field(Some(node), "content")).to_vec(),
            Some("complexField") => array(field(Some(node), "fieldResult")).to_vec(),
            Some("inlineSdt" | "insertion" | "moveTo") => {
                shown_runs(array(field(Some(node), "content")))
            }
            Some("rawXml") => array(field(Some(node), "shown")).to_vec(),
            _ => Vec::new(),
        })
        .collect()
}

fn tracked_mark(info: &Value, kind: &str, is_move_pair: bool) -> Mark {
    mark(
        kind,
        ordered_object([
            ("revisionId", nullish(field(Some(info), "id"))),
            ("author", nullish(field(Some(info), "author"))),
            ("date", nullish(field(Some(info), "date"))),
            ("isMovePair", Value::Bool(is_move_pair)),
        ]),
    )
}

fn tracked_to_units(
    content: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let content_type = string(field(Some(content), "type")).unwrap_or_default();
    let kind = if matches!(content_type, "insertion" | "moveTo") {
        "insertion"
    } else {
        "deletion"
    };
    let marker = tracked_mark(
        field(Some(content), "info").unwrap_or(&Value::Null),
        kind,
        matches!(content_type, "moveFrom" | "moveTo"),
    );
    array(field(Some(content), "content"))
        .iter()
        .flat_map(|child| {
            inline_to_units(
                child,
                style_formatting,
                styles,
                std::slice::from_ref(&marker),
                source,
            )
        })
        .collect()
}

fn sdt_properties_attrs(properties: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    map_from_value(json!({
        "sdtType": nullish(field(Some(properties), "sdtType")),
        "id": nullish(field(Some(properties), "id")),
        "alias": nullish(field(Some(properties), "alias")),
        "tag": nullish(field(Some(properties), "tag")),
        "lock": nullish(field(Some(properties), "lock")),
        "placeholder": nullish(field(Some(properties), "placeholder")),
        "showingPlaceholder": boolean(field(Some(properties), "showingPlaceholder")).unwrap_or(false),
        "dateFormat": nullish(field(Some(properties), "dateFormat")),
        "listItems": field(Some(properties), "listItems").map(|value| source_json(value, source)),
        "checked": nullish(field(Some(properties), "checked")),
        "dataBinding": field(Some(properties), "dataBinding").map(|value| source_json(value, source)),
        "rawPropertiesXml": nullish(field(Some(properties), "rawPropertiesXml")),
        "rawEndPropertiesXml": nullish(field(Some(properties), "rawEndPropertiesXml"))
    }))
}

fn sdt_payload(
    sdt: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
) -> JsonObject {
    let mut content = Vec::new();
    let append = |content: &mut Vec<Value>, unit: InlineUnit| match unit.content {
        UnitContent::Text(text) if text == "\t" => {
            content.push(json!({ "kind": "tab", "attrs": value_from_map(&unit.attrs) }));
        }
        UnitContent::Text(text) => {
            if let Some(previous) = content.last_mut()
                && string(field(Some(&*previous), "kind")) == Some("text")
                && field(Some(&*previous), "attrs") == Some(&value_from_map(&unit.attrs))
            {
                let previous_text = previous
                    .as_object_mut()
                    .and_then(|value| value.get_mut("text"))
                    .and_then(|value| value.as_str())
                    .unwrap_or_default()
                    .to_owned();
                previous.as_object_mut().unwrap().insert(
                    "text".to_owned(),
                    Value::String(format!("{previous_text}{text}")),
                );
            } else {
                content.push(json!({
                    "kind": "text",
                    "text": text,
                    "attrs": value_from_map(&unit.attrs)
                }));
            }
        }
        UnitContent::Embed { kind, payload } => {
            content.push(json!({
                "kind": kind,
                "payload": value_from_map(&payload),
                "attrs": value_from_map(&unit.attrs)
            }));
        }
    };
    for child in array(field(Some(sdt), "content")) {
        for unit in inline_to_units(child, style_formatting, styles, &[], source) {
            append(&mut content, unit);
        }
    }
    let properties = field(Some(sdt), "properties").unwrap_or(&Value::Null);
    let mut result = sdt_properties_attrs(properties, source);
    result.insert(
        "propertiesJson".to_owned(),
        Value::String(source_json(properties, source)),
    );
    result.insert("content".to_owned(), Value::Array(content));
    result
}

fn paragraph_style_formatting(
    paragraph: &Value,
    styles: &StyleResolver,
    extra: Option<&Value>,
) -> Option<Value> {
    let style_id = string(field(field(Some(paragraph), "formatting"), "styleId"));
    let style = styles
        .enabled
        .then(|| styles.resolve_paragraph_style(style_id).1)
        .flatten();
    merge_text_formatting(style.as_ref(), extra)
}

/// Note number marks carry no story unit, so the run boundary cache is the only
/// place a saved paragraph can learn they were there.
fn note_ref_mark_types(run: &Value) -> Vec<Value> {
    array(field(Some(run), "content"))
        .iter()
        .filter_map(
            |content| match string(field(Some(content), "type")).unwrap_or_default() {
                "footnoteRefMark" => Some(Value::String("footnote".to_owned())),
                "endnoteRefMark" => Some(Value::String("endnote".to_owned())),
                _ => None,
            },
        )
        .collect()
}

fn units_text(units: &[InlineUnit]) -> String {
    units
        .iter()
        .map(|unit| match &unit.content {
            UnitContent::Text(text) => text.clone(),
            UnitContent::Embed { payload, .. } => payload
                .get("footnoteRefId")
                .or_else(|| payload.get("endnoteRefId"))
                .map(js_string)
                .unwrap_or_default(),
        })
        .collect()
}

/// `w:br w:type="page"|"column"`, which the story carries as a block embed
/// beside the paragraph instead of as an inline unit.
fn flow_break_type(content: &Value) -> Option<&'static str> {
    if string(field(Some(content), "type")) != Some("break") {
        return None;
    }
    match string(field(Some(content), "breakType")) {
        Some("page") => Some("page"),
        Some("column") => Some("column"),
        _ => None,
    }
}

/// Where a run's flow breaks sit in its text. They occupy no story unit, so
/// the save projection rebuilds them from these offsets.
fn flow_break_offsets(run: &Value, source: &BTreeMap<String, String>) -> Vec<Value> {
    let contents = array(field(Some(run), "content"));
    if !contents.iter().any(|item| flow_break_type(item).is_some()) {
        return Vec::new();
    }
    let mut breaks = Vec::new();
    let mut offset = 0usize;
    for content in contents {
        if let Some(kind) = flow_break_type(content) {
            breaks.push(json!({ "offset": offset, "type": kind }));
            continue;
        }
        offset += units_text(&run_content_to_units(content, &[], source))
            .encode_utf16()
            .count();
    }
    breaks
}

fn run_boundary(
    run: &Value,
    units: &[InlineUnit],
    source: &BTreeMap<String, String>,
) -> Option<Value> {
    if units
        .iter()
        .any(|unit| matches!(&unit.content, UnitContent::Embed { kind, .. } if kind != "noteRef"))
    {
        return None;
    }
    let keys: Vec<_> = units.iter().map(|unit| marks_key(&unit.marks)).collect();
    if keys
        .first()
        .is_some_and(|first| keys.iter().any(|key| key != first))
    {
        return None;
    }
    let note_marks = note_ref_mark_types(run);
    let breaks = flow_break_offsets(run, source);
    let text = units_text(units);
    let empty = text.is_empty();
    let mut boundary = Map::new();
    boundary.insert("text".to_owned(), Value::String(text));
    if !note_marks.is_empty() {
        boundary.insert("noteMarks".to_owned(), Value::Array(note_marks));
    }
    if !breaks.is_empty() {
        boundary.insert("breaks".to_owned(), Value::Array(breaks));
    }
    if let Some(key) = keys.first() {
        boundary.insert("marksKey".to_owned(), Value::String(key.clone()));
    }
    // Only a run without text restores from cached formatting; others use live marks.
    if let Some(formatting) = field(Some(run), "formatting")
        && empty
    {
        boundary.insert("formatting".to_owned(), formatting.clone());
    }
    if let Some(changes) = field(Some(run), "propertyChanges") {
        boundary.insert("propertyChanges".to_owned(), changes.clone());
    }
    Some(Value::Object(boundary))
}

fn resolved_text_formatting(formatting: Option<&Value>, styles: &StyleResolver) -> Option<Value> {
    let style = formatting
        .and_then(|formatting| string(field(Some(formatting), "styleId")))
        .and_then(|style_id| styles.resolve_run_style(Some(style_id)));
    merge_text_formatting(style.as_ref(), formatting)
}

fn paragraph_attrs(
    paragraph: &Value,
    styles: &StyleResolver,
    units: &[InlineUnit],
    unit_counts: &[usize],
    run_boundaries: Option<Vec<Value>>,
) -> JsonObject {
    let formatting = field(Some(paragraph), "formatting");
    let style_id = string(field(formatting, "styleId"));
    let list = field(Some(paragraph), "listRendering");
    let direct_value = field(formatting, "indentFirstLine");
    let direct_nonzero = direct_value.filter(|value| number(Some(value)) != Some(0.0));
    let list_value = field(list, "indentFirstLine");
    let (selected_first, selected_hanging) = if let Some(value) = direct_nonzero {
        (Some(value), field(formatting, "hangingIndent"))
    } else if let Some(value) = list_value {
        (Some(value), field(list, "hangingIndent"))
    } else if let Some(value) = direct_value {
        (Some(value), field(formatting, "hangingIndent"))
    } else {
        (None, None)
    };
    let mut attrs = map_from_value(json!({
        "paraId": nullish(field(Some(paragraph), "paraId")),
        "textId": nullish(field(Some(paragraph), "textId")),
        "styleId": style_id,
        "numPr": nullish(field(formatting, "numPr")),
        "numPrFromStyle": nullish(field(formatting, "numPrFromStyle")),
        "listNumFmt": nullish(field(list, "numFmt")),
        "listIsBullet": nullish(field(list, "isBullet")),
        "listMarker": nullish(field(list, "marker")),
        "listMarkerHidden": truthy(field(list, "markerHidden")).then(|| field(list, "markerHidden").cloned()).flatten(),
        "listMarkerFontFamily": string(field(list, "markerFontFamily")).filter(|value| !value.is_empty()),
        "listMarkerFontSize": number(field(list, "markerFontSize")).filter(|value| *value != 0.0),
        "listMarkerBold": nullish(field(list, "markerBold")),
        "listMarkerItalic": nullish(field(list, "markerItalic")),
        "listMarkerColor": nullish(field(list, "markerColor")),
        "listMarkerSuffix": string(field(list, "markerSuffix")).filter(|value| !value.is_empty()),
        "listLevelNumFmts": truthy(field(list, "levelNumFmts")).then(|| field(list, "levelNumFmts").cloned()).flatten(),
        "listAbstractNumId": nullish(field(list, "abstractNumId")),
        "listStartOverride": nullish(field(list, "startOverride")),
        "_originalFormatting": nullish(formatting)
    }));
    if styles.enabled {
        let (style_ppr, resolved_run) = styles.resolve_paragraph_style(style_id);
        let style_ppr_ref = style_ppr.as_ref();
        for key in [
            "alignment",
            "spaceBefore",
            "spaceAfter",
            "spaceBeforeLines",
            "spaceAfterLines",
            "beforeAutospacing",
            "afterAutospacing",
            "lineSpacing",
            "lineSpacingRule",
            "indentRight",
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
        ] {
            attrs.insert(
                key.to_owned(),
                field(formatting, key)
                    .or_else(|| field(style_ppr_ref, key))
                    .cloned()
                    .unwrap_or(Value::Null),
            );
        }
        attrs.insert(
            "spacingExplicit".to_owned(),
            truthy(field(formatting, "spacingExplicit"))
                .then(|| field(formatting, "spacingExplicit").cloned())
                .flatten()
                .unwrap_or(Value::Null),
        );
        let numbering_removed = number(field(field(formatting, "numPr"), "numId")) == Some(0.0)
            && field(style_ppr_ref, "numPr").is_some()
            && number(field(field(style_ppr_ref, "numPr"), "numId")) != Some(0.0);
        attrs.insert(
            "indentLeft".to_owned(),
            field(formatting, "indentLeft")
                .or_else(|| field(list, "indentLeft"))
                .or_else(|| field(style_ppr_ref, "indentLeft"))
                .cloned()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "indentFirstLine".to_owned(),
            if selected_first.is_some() {
                selected_first
            } else if numbering_removed {
                None
            } else {
                field(style_ppr_ref, "indentFirstLine")
            }
            .cloned()
            .unwrap_or(Value::Null),
        );
        attrs.insert(
            "hangingIndent".to_owned(),
            if selected_first.is_some() {
                selected_hanging
            } else if numbering_removed {
                None
            } else {
                field(style_ppr_ref, "hangingIndent")
            }
            .cloned()
            .unwrap_or(Value::Bool(false)),
        );
        let default_character = styles
            .default_style("character")
            .and_then(|style| field(Some(style), "rPr"));
        let style_rpr = if default_character.is_some() {
            merge_text_formatting(resolved_run.as_ref(), default_character)
        } else {
            resolved_run
        };
        let direct = resolved_text_formatting(field(formatting, "runProperties"), styles);
        attrs.insert(
            "defaultTextFormatting".to_owned(),
            merge_text_formatting(style_rpr.as_ref(), direct.as_ref()).unwrap_or(Value::Null),
        );
        if field(formatting, "numPr").is_none()
            && field(style_ppr_ref, "numPr").is_some()
            && number(field(field(style_ppr_ref, "numPr"), "numId")) != Some(0.0)
        {
            let num_pr = field(style_ppr_ref, "numPr").unwrap().clone();
            attrs.insert("numPr".to_owned(), num_pr.clone());
            attrs.insert("numPrFromStyle".to_owned(), num_pr);
        }
    } else {
        for key in [
            "alignment",
            "spaceBefore",
            "spaceAfter",
            "spaceBeforeLines",
            "spaceAfterLines",
            "beforeAutospacing",
            "afterAutospacing",
            "lineSpacing",
            "lineSpacingRule",
            "indentRight",
            "borders",
            "shading",
            "tabs",
            "pageBreakBefore",
            "keepNext",
            "keepLines",
            "widowControl",
            "snapToGrid",
            "autoSpaceDE",
            "autoSpaceDN",
            "outlineLevel",
            "bidi",
        ] {
            attrs.insert(
                key.to_owned(),
                field(formatting, key).cloned().unwrap_or(Value::Null),
            );
        }
        attrs.insert(
            "spacingExplicit".to_owned(),
            truthy(field(formatting, "spacingExplicit"))
                .then(|| field(formatting, "spacingExplicit").cloned())
                .flatten()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "indentLeft".to_owned(),
            field(formatting, "indentLeft")
                .or_else(|| field(list, "indentLeft"))
                .cloned()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "indentFirstLine".to_owned(),
            selected_first.cloned().unwrap_or(Value::Null),
        );
        attrs.insert(
            "hangingIndent".to_owned(),
            selected_hanging.cloned().unwrap_or(Value::Bool(false)),
        );
        attrs.insert(
            "defaultTextFormatting".to_owned(),
            field(formatting, "runProperties")
                .cloned()
                .unwrap_or(Value::Null),
        );
    }
    if let Some(section) = field(Some(paragraph), "sectionProperties") {
        attrs.insert("_sectionProperties".to_owned(), section.clone());
        if let Some(start @ ("nextPage" | "continuous" | "oddPage" | "evenPage" | "nextColumn")) =
            string(field(Some(section), "sectionStart"))
        {
            attrs.insert(
                "sectionBreakType".to_owned(),
                Value::String(start.to_owned()),
            );
        }
    }
    if truthy(field(Some(paragraph), "renderedPageBreakBefore")) {
        attrs.insert("renderedPageBreakBefore".to_owned(), Value::Bool(true));
    }
    for (source, target) in [("pPrIns", "pPrIns"), ("pPrDel", "pPrDel")] {
        if let Some(info) = field(Some(paragraph), source) {
            attrs.insert(
                target.to_owned(),
                json!({
                    "revisionId": nullish(field(Some(info), "id")),
                    "author": nullish(field(Some(info), "author")),
                    "date": nullish(field(Some(info), "date"))
                }),
            );
        }
    }
    if !array(field(Some(paragraph), "propertyChanges")).is_empty() {
        attrs.insert(
            "pPrChange".to_owned(),
            field(Some(paragraph), "propertyChanges").unwrap().clone(),
        );
    }
    let mut bookmarks = Vec::new();
    let mut unit_index = 0usize;
    let mut pm_offset = 0u32;
    // A bookmark ahead of some of the paragraph's leading breaks records how
    // many follow it, so the save writes it back before them.
    let leading_items: Vec<usize> = paragraph_flow_breaks(paragraph)
        .0
        .iter()
        .map(|flow| flow.item)
        .collect();
    let breaks_after = |bookmark: &mut Value, content_index: usize| {
        let after = leading_items
            .iter()
            .filter(|&&item| item > content_index)
            .count();
        if after > 0 {
            bookmark["breaksAfter"] = json!(after);
        }
    };
    for (content_index, content) in array(field(Some(paragraph), "content")).iter().enumerate() {
        match string(field(Some(content), "type")).unwrap_or_default() {
            "bookmarkStart" => {
                let mut bookmark = json!({
                    "id": nullish(field(Some(content), "id")),
                    "name": nullish(field(Some(content), "name")),
                    "kind": "start",
                    "offset": pm_offset
                });
                for key in ["colFirst", "colLast"] {
                    if let Some(value) = field(Some(content), key) {
                        bookmark
                            .as_object_mut()
                            .unwrap()
                            .insert(key.to_owned(), value.clone());
                    }
                }
                breaks_after(&mut bookmark, content_index);
                bookmarks.push(bookmark);
            }
            "bookmarkEnd" => {
                let mut bookmark = json!({
                    "id": nullish(field(Some(content), "id")),
                    "kind": "end",
                    "offset": pm_offset
                });
                breaks_after(&mut bookmark, content_index);
                bookmarks.push(bookmark);
            }
            _ => {
                for _ in 0..unit_counts.get(content_index).copied().unwrap_or_default() {
                    pm_offset += units.get(unit_index).map(|unit| unit.pm_size).unwrap_or(0);
                    unit_index += 1;
                }
            }
        }
    }
    if !bookmarks.is_empty() {
        attrs.insert("bookmarks".to_owned(), Value::Array(bookmarks));
    }
    if let Some(boundaries) = run_boundaries.filter(|boundaries| !boundaries.is_empty()) {
        attrs.insert(
            "_originalRunBoundaries".to_owned(),
            Value::Array(boundaries),
        );
    }
    attrs
}

fn para_attrs_to_ppr(attrs: JsonObject) -> JsonObject {
    attrs
        .into_iter()
        .filter(|(key, value)| {
            ![
                "paraId",
                "textId",
                "renderedPageBreakBefore",
                "numPrFromStyle",
            ]
            .contains(&key.as_str())
                && !value.is_null()
        })
        .map(|(key, value)| {
            (
                match key.as_str() {
                    "styleId" => "pStyle".to_owned(),
                    "_sectionProperties" => "sectPr".to_owned(),
                    _ => key,
                },
                drop_nulls(value),
            )
        })
        .collect()
}

/// Comment ranges open at the lowering point, each with the story it opened in.
#[derive(Default)]
struct OpenComments(Vec<(String, usize)>);

impl OpenComments {
    fn open(&mut self, id: String, story: usize) {
        if !self.0.iter().any(|(open, _)| *open == id) {
            self.0.push((id, story));
        }
    }

    fn close(&mut self, id: &str) {
        self.0.retain(|(open, _)| open != id);
    }

    /// Opens or closes the range a `commentRangeStart`/`commentRangeEnd` marks.
    fn mark(&mut self, marker: &Value, story: usize) {
        let Some(id) = field(Some(marker), "id") else {
            return;
        };
        match string(field(Some(marker), "type")) {
            Some("commentRangeStart") => self.open(js_string(id), story),
            Some("commentRangeEnd") => self.close(&js_string(id)),
            _ => {}
        }
    }

    /// The comments a unit of `story` falls inside: a range covers only the story it opened in.
    fn covering(&self, story: usize) -> Vec<String> {
        self.0
            .iter()
            .filter(|(_, home)| *home == story)
            .map(|(id, _)| id.clone())
            .collect()
    }
}

fn paragraph_units(
    paragraph: &Value,
    styles: &StyleResolver,
    extra_run_formatting: Option<&Value>,
    source: &BTreeMap<String, String>,
    comments: &mut OpenComments,
    story: usize,
) -> (Vec<InlineUnit>, JsonObject) {
    let mut units = Vec::new();
    let mut boundaries = Some(Vec::new());
    let mut unit_counts = Vec::new();
    let (_, _, inline_breaks) = paragraph_flow_plan(paragraph);
    let style_formatting = paragraph_style_formatting(paragraph, styles, extra_run_formatting);
    for (content_index, content) in array(field(Some(paragraph), "content")).iter().enumerate() {
        let start = units.len();
        let mut break_index = 0;
        let mut keep_inline = |unit: &InlineUnit| {
            if matches!(&unit.content, UnitContent::Embed { kind, .. } if kind == "pageBreak" || kind == "columnBreak")
            {
                let keep = inline_breaks[content_index]
                    .get(break_index)
                    .copied()
                    .unwrap_or(false);
                break_index += 1;
                keep
            } else {
                true
            }
        };
        match string(field(Some(content), "type")).unwrap_or_default() {
            "commentRangeStart" | "commentRangeEnd" => comments.mark(content, story),
            "run" => {
                let run_units: Vec<_> =
                    run_to_units(content, style_formatting.as_ref(), styles, &[], source)
                        .into_iter()
                        .filter(&mut keep_inline)
                        .collect();
                if let Some(run_boundaries) = &mut boundaries {
                    if let Some(boundary) = run_boundary(content, &run_units, source) {
                        run_boundaries.push(boundary);
                    } else {
                        boundaries = None;
                    }
                }
                units.extend(run_units);
            }
            "hyperlink" => {
                boundaries = None;
                units.extend(hyperlink_to_units(
                    content,
                    style_formatting.as_ref(),
                    styles,
                    &[],
                    source,
                ));
            }
            "simpleField" | "complexField" => {
                boundaries = None;
                units.extend(field_to_units(
                    content,
                    style_formatting.as_ref(),
                    styles,
                    source,
                    unit_counts.len(),
                ));
            }
            "inlineSdt" => {
                boundaries = None;
                units.push(embed_unit(
                    "sdt",
                    sdt_payload(content, style_formatting.as_ref(), styles, source),
                    &[],
                    2,
                ));
            }
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                boundaries = None;
                units.extend(tracked_to_units(
                    content,
                    style_formatting.as_ref(),
                    styles,
                    source,
                ));
            }
            "mathEquation" => {
                boundaries = None;
                units.push(embed_unit("math", math_payload(content), &[], 1));
            }
            "bookmarkStart" | "bookmarkEnd" | "rawXml" => {}
            _ => boundaries = None,
        }
        if string(field(Some(content), "type")) != Some("run") {
            let mut content_units = units.split_off(start);
            content_units.retain(keep_inline);
            units.extend(content_units);
        }
        for unit in &mut units[start..] {
            unit.comments = comments.covering(story);
        }
        unit_counts.push(units.len() - start);
    }
    // The cache is kept only for what merging equal-formatted runs would lose.
    let boundaries = boundaries.filter(|boundaries| {
        boundaries.iter().any(|boundary| {
            boundary.get("propertyChanges").is_some()
                || boundary.get("noteMarks").is_some()
                || boundary.get("breaks").is_some()
                || boundary.get("text").and_then(Value::as_str) == Some("")
        })
    });
    let attrs = paragraph_attrs(paragraph, styles, &units, &unit_counts, boundaries);
    (units, para_attrs_to_ppr(attrs))
}

/// A page or column break (with the tracked change around it), or visible
/// content, in the order Word reads a paragraph.
type FlowToken = (&'static str, Option<Mark>);

fn run_tokens(run: &Value, marker: Option<&Mark>, tokens: &mut Vec<FlowToken>) {
    for content in array(field(Some(run), "content")) {
        if let Some(kind) = flow_break_type(content) {
            tokens.push((
                if kind == "column" {
                    "columnBreak"
                } else {
                    "pageBreak"
                },
                marker.cloned(),
            ));
        } else if string(field(Some(content), "type")) == Some("commentReference") {
            tokens.push(("reference", None));
        } else if string(field(Some(content), "type")) != Some("text")
            || !string(field(Some(content), "text"))
                .unwrap_or_default()
                .is_empty()
        {
            tokens.push(("visible", None));
        }
    }
}

fn inline_tokens(content: &[Value], tokens: &mut Vec<FlowToken>) {
    nested_tokens(content, None, tokens);
}

fn nested_tokens(content: &[Value], marker: Option<&Mark>, tokens: &mut Vec<FlowToken>) {
    for item in content {
        match string(field(Some(item), "type")).unwrap_or_default() {
            "run" => run_tokens(item, marker, tokens),
            "hyperlink" => nested_tokens(
                array(
                    field(Some(item), "structuredChildren")
                        .or_else(|| field(Some(item), "children")),
                ),
                marker,
                tokens,
            ),
            "simpleField" => {
                tokens.push(("visible", None));
                nested_tokens(
                    array(
                        field(field(Some(item), "structuredResult"), "inline")
                            .or_else(|| field(Some(item), "content")),
                    ),
                    marker,
                    tokens,
                );
            }
            "bookmarkStart" | "bookmarkEnd" => tokens.push(("mark", None)),
            "complexField" => {
                for (structured, legacy) in [
                    ("structuredCode", "fieldCode"),
                    ("structuredResult", "fieldResult"),
                ] {
                    nested_tokens(
                        array(
                            field(field(Some(item), structured), "inline")
                                .or_else(|| field(Some(item), legacy)),
                        ),
                        marker,
                        tokens,
                    );
                }
            }
            "inlineSdt" => nested_tokens(array(field(Some(item), "content")), marker, tokens),
            kind @ ("insertion" | "deletion" | "moveFrom" | "moveTo") => {
                let own = tracked_mark(
                    field(Some(item), "info").unwrap_or(&Value::Null),
                    if matches!(kind, "insertion" | "moveTo") {
                        "insertion"
                    } else {
                        "deletion"
                    },
                    matches!(kind, "moveFrom" | "moveTo"),
                );
                nested_tokens(array(field(Some(item), "content")), Some(&own), tokens);
            }
            "mathEquation" => tokens.push(("visible", None)),
            _ => {}
        }
    }
}

fn exposed_flow_breaks(item: &Value, project_field: bool) -> usize {
    match string(field(Some(item), "type")) {
        Some("run") => array(field(Some(item), "content"))
            .iter()
            .filter(|content| flow_break_type(content).is_some())
            .count(),
        Some("hyperlink") => {
            array(field(Some(item), "structuredChildren").or_else(|| field(Some(item), "children")))
                .iter()
                .map(|child| exposed_flow_breaks(child, false))
                .sum()
        }
        Some("insertion" | "deletion" | "moveFrom" | "moveTo") => {
            array(field(Some(item), "content"))
                .iter()
                .map(|child| exposed_flow_breaks(child, false))
                .sum()
        }
        Some("complexField")
            if project_field
                && !numeric_field_instruction(
                    string(field(Some(item), "instruction")).unwrap_or_default(),
                ) =>
        {
            ["structuredCode", "structuredResult"]
                .iter()
                .flat_map(|key| array(field(field(Some(item), key), "inline")))
                .filter(|child| string(field(Some(child), "type")) == Some("hyperlink"))
                .map(|child| exposed_flow_breaks(child, false))
                .sum()
        }
        _ => 0,
    }
}

/// A break unit the seed places around its paragraph, `leading` when Word
/// wrote it before the paragraph's text (which keeps the paragraph's
/// space-before), from the paragraph's content item `item`.
struct FlowBreak {
    kind: &'static str,
    leading: bool,
    trailing: bool,
    marker: Option<Mark>,
    item: usize,
}

impl FlowBreak {
    fn unit(self) -> InlineUnit {
        let mut payload = JsonObject::new();
        if self.leading {
            payload.insert("leading".to_owned(), Value::Bool(true));
        }
        if self.trailing {
            payload.insert("trailing".to_owned(), Value::Bool(true));
        }
        embed_unit(self.kind, payload, self.marker.as_slice(), 1)
    }
}

/// The breaks that open a paragraph's slot (before its text) and those that
/// follow its pilcrow (after its text). A paragraph without text keeps its
/// breaks after its pilcrow, except those up to its last column break.
fn paragraph_flow_breaks(paragraph: &Value) -> (Vec<FlowBreak>, Vec<FlowBreak>) {
    let (leading, trailing, _) = paragraph_flow_plan(paragraph);
    (leading, trailing)
}

fn paragraph_flow_plan(paragraph: &Value) -> (Vec<FlowBreak>, Vec<FlowBreak>, Vec<Vec<bool>>) {
    let mut tokens = Vec::new();
    let mut items = Vec::new();
    for (index, item) in array(field(Some(paragraph), "content")).iter().enumerate() {
        inline_tokens(std::slice::from_ref(item), &mut tokens);
        items.resize(tokens.len(), index);
    }
    let is_break = |index: &usize| matches!(tokens[*index].0, "pageBreak" | "columnBreak");
    // Breaks lead the content that follows them: the text, or without text
    // the last bookmark. Without either, those up to the last column break lead.
    let last_break = (0..tokens.len()).rfind(is_break);
    let text = tokens.iter().enumerate().position(|(index, (kind, _))| {
        *kind == "visible" || (*kind == "reference" && last_break.is_none_or(|last| index > last))
    });
    let content = text.or_else(|| tokens.iter().rposition(|(kind, _)| *kind == "mark"));
    let split = content.unwrap_or_else(|| {
        tokens
            .iter()
            .rposition(|(kind, _)| *kind == "columnBreak")
            .map_or(0, |index| index + 1)
    });
    let last_text = tokens.iter().rposition(|(kind, _)| *kind == "visible");
    let mut inline_breaks = vec![Vec::new(); array(field(Some(paragraph), "content")).len()];
    let exposed: Vec<_> = array(field(Some(paragraph), "content"))
        .iter()
        .map(|item| exposed_flow_breaks(item, true))
        .collect();
    let mut counts = vec![0; exposed.len()];
    for index in (0..tokens.len()).filter(is_break) {
        counts[items[index]] += 1;
    }
    let inline: Vec<_> = (0..tokens.len())
        .map(|index| {
            is_break(&index)
                && counts[items[index]] == exposed[items[index]]
                && text.is_some_and(|first| first < index)
                && last_text.is_some_and(|last| index < last)
        })
        .collect();
    for index in (0..tokens.len()).filter(is_break) {
        inline_breaks[items[index]].push(inline[index]);
    }
    let breaks = |range: std::ops::Range<usize>| {
        range
            .filter(|index| is_break(index) && !inline[*index])
            .map(|index| FlowBreak {
                kind: tokens[index].0,
                leading: content.is_some_and(|content| index < content),
                trailing: index >= split,
                marker: tokens[index].1.clone(),
                item: items[index],
            })
            .collect()
    };
    (breaks(0..split), breaks(split..tokens.len()), inline_breaks)
}

fn modifier(value: &str) -> f64 {
    let prefix: String = value
        .chars()
        .take_while(|character| character.is_ascii_hexdigit())
        .collect();
    u8::from_str_radix(&prefix, 16)
        .map(|value| f64::from(value) / 255.0)
        .unwrap_or(1.0)
}

fn rgb_channels(value: &str) -> [u8; 3] {
    let mut normalized = value.trim_start_matches('#').to_owned();
    while normalized.len() < 6 {
        normalized.insert(0, '0');
    }
    normalized.truncate(6);
    [
        u8::from_str_radix(&normalized[0..2], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[2..4], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[4..6], 16).unwrap_or(0),
    ]
}

fn resolve_color_to_hex(color: Option<&Value>, theme: Option<&Value>) -> Option<String> {
    let color = object(color)?;
    if truthy(color.get("auto")) {
        return None;
    }
    if let Some(theme_color) = string(color.get("themeColor"))
        && let Some(theme) = theme
    {
        let slot = match theme_color {
            "dark1" | "text1" | "tx1" => "dk1",
            "light1" | "background1" | "bg1" => "lt1",
            "dark2" | "text2" | "tx2" => "dk2",
            "light2" | "background2" | "bg2" => "lt2",
            "hyperlink" => "hlink",
            "followedHyperlink" => "folHlink",
            value => value,
        };
        let known = [
            "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
            "accent6", "hlink", "folHlink",
        ];
        let mut hex = if known.contains(&slot) {
            string(field(field(Some(theme), "colorScheme"), slot))
                .or_else(|| string(color.get("rgb")))
                .unwrap_or("000000")
                .to_owned()
        } else {
            string(color.get("rgb")).unwrap_or("000000").to_owned()
        };
        let mut channels = rgb_channels(&hex);
        if let Some(tint) = string(color.get("themeTint")) {
            let tint = modifier(tint);
            channels = channels
                .map(|channel| (f64::from(channel) * tint + 255.0 * (1.0 - tint)).round() as u8);
        } else if let Some(shade) = string(color.get("themeShade")) {
            let shade = modifier(shade);
            channels = channels.map(|channel| (f64::from(channel) * shade).round() as u8);
        }
        hex = format!("{:02X}{:02X}{:02X}", channels[0], channels[1], channels[2]);
        return Some(hex);
    }
    string(color.get("rgb"))
        .filter(|value| *value != "auto")
        .map(|value| value.trim_start_matches('#').to_ascii_uppercase())
}

fn calculate_row_spans(table: &Value) -> BTreeMap<(usize, usize), (usize, bool)> {
    let mut result = BTreeMap::new();
    let mut active = BTreeMap::<usize, usize>::new();
    for (row_index, row) in array(field(Some(table), "rows")).iter().enumerate() {
        let mut column = 0usize;
        let cells: Vec<_> = array(field(Some(row), "cells"))
            .iter()
            .map(|cell| {
                let current = column;
                column += number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                    as usize;
                (
                    current,
                    string(field(field(Some(cell), "formatting"), "vMerge")),
                )
            })
            .collect();
        let empty = !cells.is_empty()
            && cells
                .iter()
                .all(|(column, merge)| *merge == Some("continue") && active.contains_key(column));
        if empty {
            for (column, _) in cells {
                active.remove(&column);
                result.insert((row_index, column), (1, false));
            }
            continue;
        }
        for (column, merge) in cells {
            match merge {
                Some("restart") => {
                    active.insert(column, row_index);
                    result.insert((row_index, column), (1, false));
                }
                Some("continue") => {
                    if let Some(start) = active.get(&column).copied() {
                        if let Some(owner) = result.get_mut(&(start, column)) {
                            owner.0 += 1;
                        }
                        result.insert((row_index, column), (1, true));
                    } else {
                        result.insert((row_index, column), (1, false));
                    }
                }
                _ => {
                    active.remove(&column);
                    result.insert((row_index, column), (1, false));
                }
            }
        }
    }
    result
}

fn revision_attrs(info: &Value) -> Value {
    json!({
        "revisionId": nullish(field(Some(info), "id")),
        "author": nullish(field(Some(info), "author")),
        "date": nullish(field(Some(info), "date"))
    })
}

/// Maps each physical cell edge to the border side that feeds it: the matching
/// outer side where the cell sits on the boundary, `insideH`/`insideV` within.
pub(crate) fn border_side_sources(
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
) -> [(&'static str, &'static str); 4] {
    [
        ("top", if first_row { "top" } else { "insideH" }),
        ("bottom", if last_row { "bottom" } else { "insideH" }),
        ("left", if first_column { "left" } else { "insideV" }),
        ("right", if last_column { "right" } else { "insideV" }),
    ]
}

fn cell_borders(
    formatting: Option<&Value>,
    table_borders: Option<&Value>,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
) -> Option<Value> {
    let inherited = object(table_borders).map(|borders| {
        Value::Object(
            border_side_sources(first_row, last_row, first_column, last_column)
                .into_iter()
                .map(|(edge, source)| (edge.to_owned(), nullish(borders.get(source))))
                .collect(),
        )
    });
    let direct = field(formatting, "borders");
    if inherited.is_none() && direct.is_none() {
        None
    } else {
        merge_plain(inherited.as_ref(), direct)
    }
}

struct CellOptions<'a> {
    is_header: bool,
    rowspan: usize,
    grid_width: Option<f64>,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    table_borders: Option<&'a Value>,
    default_margins: Option<&'a Value>,
    theme: Option<&'a Value>,
    table_bidi: bool,
}

fn structural_attrs(attrs: JsonObject, skipped: &[&str]) -> JsonObject {
    attrs
        .into_iter()
        .filter(|(key, value)| !skipped.contains(&key.as_str()) && !value.is_null())
        .map(|(key, value)| (key, drop_nulls(value)))
        .collect()
}

fn project_cell(cell: &Value, options: CellOptions<'_>) -> ProjectedCell {
    let formatting = field(Some(cell), "formatting");
    let background =
        resolve_color_to_hex(field(field(formatting, "shading"), "fill"), options.theme);
    let width = field(field(formatting, "width"), "value")
        .cloned()
        .or_else(|| options.grid_width.map(|value| json!(value)));
    let width_type = field(field(formatting, "width"), "type")
        .cloned()
        .or_else(|| options.grid_width.map(|_| Value::String("pct".to_owned())));
    let margins = if let Some(margins) = field(formatting, "margins") {
        Some(json!({
            "top": nullish(field(field(Some(margins), "top"), "value")),
            "bottom": nullish(field(field(Some(margins), "bottom"), "value")),
            "left": nullish(
                field(field(Some(margins), "left"), "value").or_else(|| {
                    field(
                        field(Some(margins), if options.table_bidi { "end" } else { "start" }),
                        "value",
                    )
                })
            ),
            "right": nullish(
                field(field(Some(margins), "right"), "value").or_else(|| {
                    field(
                        field(Some(margins), if options.table_bidi { "start" } else { "end" }),
                        "value",
                    )
                })
            )
        }))
    } else {
        options.default_margins.cloned()
    };
    let mut attrs = map_from_value(json!({
        "colspan": number(field(formatting, "gridSpan")).unwrap_or(1.0),
        "rowspan": options.rowspan,
        "width": width,
        "widthType": width_type,
        "verticalAlign": nullish(field(formatting, "verticalAlign")),
        "backgroundColor": background,
        "borders": cell_borders(
            formatting,
            options.table_borders,
            options.first_row,
            options.last_row,
            options.first_column,
            options.last_column,
        ),
        "margins": margins,
        "textDirection": nullish(field(formatting, "textDirection")),
        "noWrap": boolean(field(formatting, "noWrap")).unwrap_or(false),
        "_originalFormatting": nullish(formatting),
        "_originalResolvedFill": background
    }));
    if let Some(change) = field(Some(cell), "structuralChange") {
        let info = revision_attrs(field(Some(change), "info").unwrap_or(&Value::Null));
        match string(field(Some(change), "type")).unwrap_or_default() {
            "tableCellInsertion" => {
                attrs.insert(
                    "cellMarker".to_owned(),
                    json!({ "kind": "ins", "info": info }),
                );
            }
            "tableCellDeletion" => {
                attrs.insert(
                    "cellMarker".to_owned(),
                    json!({ "kind": "del", "info": info }),
                );
            }
            "tableCellMerge" => {
                let mut marker = json!({
                    "kind": "merge",
                    "info": info,
                    "vMerge": string(field(Some(change), "vMerge")).unwrap_or("cont")
                });
                if let Some(value) =
                    string(field(Some(change), "vMergeOrig")).filter(|value| !value.is_empty())
                {
                    marker
                        .as_object_mut()
                        .unwrap()
                        .insert("vMergeOrig".to_owned(), Value::String(value.to_owned()));
                }
                attrs.insert("cellMarker".to_owned(), marker);
            }
            _ => {}
        }
    }
    if !array(field(Some(cell), "propertyChanges")).is_empty() {
        attrs.insert(
            "tcPrChange".to_owned(),
            field(Some(cell), "propertyChanges").unwrap().clone(),
        );
    }
    let mut attrs = structural_attrs(attrs, &[]);
    if options.is_header {
        attrs.insert("header".to_owned(), Value::Bool(true));
    }
    let content = array(field(Some(cell), "content"));
    ProjectedCell {
        paragraph_formatting: None,
        attrs,
        content: if content.is_empty() {
            vec![json!({ "type": "paragraph", "content": [] })]
        } else {
            content.to_vec()
        },
    }
}

struct TableStyleContext<'a> {
    column_count: usize,
    style: Option<&'a Value>,
    borders: Option<&'a Value>,
    margins: Option<&'a Value>,
    theme: Option<&'a Value>,
}

fn table_column_count(table: &Value) -> usize {
    array(field(Some(table), "rows"))
        .iter()
        .map(|row| {
            let row_formatting = field(Some(row), "formatting");
            let omitted = number(field(row_formatting, "gridBefore")).unwrap_or(0.0) as usize
                + number(field(row_formatting, "gridAfter")).unwrap_or(0.0) as usize;
            omitted
                + array(field(Some(row), "cells"))
                    .iter()
                    .map(|cell| {
                        number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                            as usize
                    })
                    .sum::<usize>()
        })
        .max()
        .unwrap_or(0)
        .max(array(field(Some(table), "columnWidths")).len())
}

fn table_cell_paragraph_formatting(
    table: &Value,
    style: Option<&Value>,
    row_index: usize,
    start_column: usize,
    end_column: usize,
    columns: usize,
) -> Option<Value> {
    let mut result = field(style, "pPr").cloned();
    let parts = array(field(style, "tblStylePr"));
    if !parts.iter().any(|part| field(Some(part), "pPr").is_some()) {
        return result;
    }
    let formatting = field(Some(table), "formatting");
    let style_formatting = field(style, "tblPr");
    let look = field(formatting, "look").or_else(|| field(style_formatting, "look"));
    let mask = string(field(look, "value"))
        .and_then(|value| u32::from_str_radix(value, 16).ok())
        .unwrap_or(if look.is_none() { 0x04a0 } else { 0 });
    let flag = |key, bit| boolean(field(look, key)).unwrap_or(mask & bit != 0);
    let first_row = flag("firstRow", 0x20);
    let last_row = flag("lastRow", 0x40);
    let first_column = flag("firstColumn", 0x80);
    let last_column = flag("lastColumn", 0x100);
    let rows = array(field(Some(table), "rows"));
    let grid_before = number(field(
        field(rows.get(row_index), "formatting"),
        "gridBefore",
    ))
    .unwrap_or(0.0) as usize;
    let start_column = start_column + grid_before;
    let end_column = end_column + grid_before;
    let at_first_row = first_row && row_index == 0;
    let at_last_row = last_row && row_index + 1 == rows.len();
    let at_first_column = first_column && start_column == 0;
    let at_last_column = last_column && end_column == columns;
    let band_size = |key| {
        number(field(formatting, key).or_else(|| field(style_formatting, key))).unwrap_or(1.0)
    };
    let row_band_size = band_size("styleRowBandSize");
    let column_band_size = band_size("styleColBandSize");
    let mut regions = Vec::new();
    if !flag("noHBand", 0x200) && !at_first_row && !at_last_row && row_band_size > 0.0 {
        let band =
            (row_index.saturating_sub(usize::from(first_row)) as f64 / row_band_size).floor();
        regions.push(if band % 2.0 == 0.0 {
            "band1Horz"
        } else {
            "band2Horz"
        });
    }
    if !flag("noVBand", 0x400) && !at_first_column && !at_last_column && column_band_size > 0.0 {
        let band = (start_column.saturating_sub(usize::from(first_column)) as f64
            / column_band_size)
            .floor();
        regions.push(if band % 2.0 == 0.0 {
            "band1Vert"
        } else {
            "band2Vert"
        });
    }
    for (region, active) in [
        ("firstCol", at_first_column),
        ("lastCol", at_last_column),
        ("firstRow", at_first_row),
        ("lastRow", at_last_row),
        ("nwCell", at_first_row && at_first_column),
        ("neCell", at_first_row && at_last_column),
        ("swCell", at_last_row && at_first_column),
        ("seCell", at_last_row && at_last_column),
    ] {
        if active {
            regions.push(region);
        }
    }
    for region in regions {
        let conditional = parts
            .iter()
            .find(|part| string(field(Some(part), "type")) == Some(region));
        result = merge_paragraph_formatting(result.as_ref(), field(conditional, "pPr"));
    }
    result
}

fn project_row(
    row: &Value,
    table: &Value,
    row_index: usize,
    row_spans: &BTreeMap<(usize, usize), (usize, bool)>,
    style_context: &TableStyleContext<'_>,
) -> ProjectedRow {
    let formatting = field(Some(row), "formatting");
    let mut attrs = map_from_value(json!({
        "height": nullish(field(field(formatting, "height"), "value")),
        "heightRule": nullish(field(formatting, "heightRule")),
        "isHeader": truthy(field(formatting, "header")),
        "_originalFormatting": nullish(formatting)
    }));
    if let Some(change) = field(Some(row), "structuralChange") {
        let value = revision_attrs(field(Some(change), "info").unwrap_or(&Value::Null));
        match string(field(Some(change), "type")).unwrap_or_default() {
            "tableRowInsertion" => {
                attrs.insert("trIns".to_owned(), value);
            }
            "tableRowDeletion" => {
                attrs.insert("trDel".to_owned(), value);
            }
            _ => {}
        }
    }
    if !array(field(Some(row), "propertyChanges")).is_empty() {
        attrs.insert(
            "trPrChange".to_owned(),
            field(Some(row), "propertyChanges").unwrap().clone(),
        );
    }
    let widths = array(field(Some(table), "columnWidths"));
    let total_width: f64 = widths.iter().filter_map(Value::as_f64).sum();
    let rows = array(field(Some(table), "rows"));
    let total_columns = if !widths.is_empty() {
        widths.len()
    } else {
        rows.iter()
            .map(|row| {
                array(field(Some(row), "cells"))
                    .iter()
                    .map(|cell| {
                        number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                            as usize
                    })
                    .sum()
            })
            .max()
            .unwrap_or(0)
    };
    let cells_source = array(field(Some(row), "cells"));
    let mut column = 0usize;
    let mut cells = Vec::new();
    for cell in cells_source {
        let colspan =
            number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0) as usize;
        let start_column = column;
        let span = row_spans.get(&(row_index, start_column));
        let grid_width = (!widths.is_empty() && total_width > 0.0).then(|| {
            let cell_width: f64 = widths
                .iter()
                .skip(start_column)
                .take(colspan)
                .filter_map(Value::as_f64)
                .sum();
            (cell_width / total_width * 100.0).round()
        });
        column += colspan;
        if span.is_some_and(|(_, skip)| *skip) {
            continue;
        }
        let mut projected = project_cell(
            cell,
            CellOptions {
                is_header: row_index == 0
                    && truthy(field(
                        field(field(Some(table), "formatting"), "look"),
                        "firstRow",
                    )),
                rowspan: span.map(|(rowspan, _)| *rowspan).unwrap_or(1),
                grid_width,
                first_row: row_index == 0,
                last_row: row_index + 1 == rows.len(),
                first_column: start_column == 0,
                last_column: column == total_columns,
                table_borders: style_context.borders,
                default_margins: style_context.margins,
                theme: style_context.theme,
                table_bidi: truthy(field(field(Some(table), "formatting"), "bidi")),
            },
        );
        projected.paragraph_formatting = table_cell_paragraph_formatting(
            table,
            style_context.style,
            row_index,
            start_column,
            column,
            style_context.column_count,
        );
        cells.push(projected);
    }
    if cells.is_empty() {
        let synthetic = if total_columns > 1 {
            json!({
                "type": "tableCell",
                "formatting": { "gridSpan": total_columns },
                "content": [{ "type": "paragraph", "content": [] }]
            })
        } else {
            json!({
                "type": "tableCell",
                "content": [{ "type": "paragraph", "content": [] }]
            })
        };
        let mut projected = project_cell(
            &synthetic,
            CellOptions {
                is_header: row_index == 0
                    && truthy(field(
                        field(field(Some(table), "formatting"), "look"),
                        "firstRow",
                    )),
                rowspan: 1,
                grid_width: (total_width > 0.0).then_some(100.0),
                first_row: row_index == 0,
                last_row: row_index + 1 == rows.len(),
                first_column: true,
                last_column: true,
                table_borders: style_context.borders,
                default_margins: style_context.margins,
                theme: style_context.theme,
                table_bidi: truthy(field(field(Some(table), "formatting"), "bidi")),
            },
        );
        projected.paragraph_formatting = table_cell_paragraph_formatting(
            table,
            style_context.style,
            row_index,
            0,
            total_columns,
            style_context.column_count,
        );
        cells.push(projected);
    }
    ProjectedRow {
        attrs: structural_attrs(attrs, &[]),
        cells,
    }
}

fn project_table(
    table: &Value,
    styles: &StyleResolver,
    theme: Option<&Value>,
    compatibility_mode: u8,
) -> ProjectedTable {
    let formatting = field(Some(table), "formatting");
    let default_style = styles.default_style("table");
    let style_id = string(field(formatting, "styleId"));
    let effective_style_id =
        style_id.or_else(|| default_style.and_then(|style| string(field(Some(style), "styleId"))));
    let table_style = effective_style_id.and_then(|id| styles.style(id));
    let borders = field(formatting, "borders")
        .or_else(|| field(field(table_style, "tblPr"), "borders"))
        .or_else(|| field(field(default_style, "tblPr"), "borders"));
    let margins = field(formatting, "cellMargins")
        .or_else(|| field(field(table_style, "tblPr"), "cellMargins"))
        .or_else(|| field(field(default_style, "tblPr"), "cellMargins"));
    let logical_left = field(
        margins,
        if truthy(field(formatting, "bidi")) {
            "end"
        } else {
            "start"
        },
    );
    let logical_right = field(
        margins,
        if truthy(field(formatting, "bidi")) {
            "start"
        } else {
            "end"
        },
    );
    let default_margins = margins.map(|margins| {
        drop_nulls(json!({
            "top": nullish(field(field(Some(margins), "top"), "value")),
            "bottom": nullish(field(field(Some(margins), "bottom"), "value")),
            "left": nullish(
                field(field(Some(margins), "left"), "value")
                    .or_else(|| field(logical_left, "value"))
            ),
            "right": nullish(
                field(field(Some(margins), "right"), "value")
                    .or_else(|| field(logical_right, "value"))
            )
        }))
    });
    let mut based_on = Vec::new();
    let mut visited = BTreeSet::new();
    let mut inherited = table_style;
    while let Some(parent_id) = inherited.and_then(|style| string(field(Some(style), "basedOn"))) {
        if based_on.len() >= 32 || !visited.insert(parent_id.to_owned()) {
            break;
        }
        based_on.insert(0, Value::String(parent_id.to_owned()));
        inherited = styles.style(parent_id);
    }
    let mut original_formatting = object(formatting).cloned().unwrap_or_default();
    original_formatting.insert(
        "styleCascade".to_owned(),
        drop_nulls(json!({
            "selectedStyleId": style_id.filter(|value| !value.is_empty()),
            "defaultStyleId": default_style
                .and_then(|style| string(field(Some(style), "styleId")))
                .filter(|value| !value.is_empty()),
            "basedOnStyleIds": (!based_on.is_empty()).then_some(based_on)
        })),
    );
    let mut attrs = map_from_value(json!({
        "styleId": style_id,
        "width": nullish(field(field(formatting, "width"), "value")),
        "widthType": nullish(field(field(formatting, "width"), "type")),
        "justification": nullish(field(formatting, "justification")),
        "columnWidths": nullish(field(Some(table), "columnWidths")),
        "tableLayout": nullish(field(formatting, "layout")),
        "floating": nullish(field(formatting, "floating")),
        "cellMargins": default_margins,
        "look": nullish(field(formatting, "look")),
        "bidi": truthy(field(formatting, "bidi")).then_some(true),
        "compatibilityMode": if compatibility_mode == 12 {
            Value::Null
        } else {
            json!(compatibility_mode as f64)
        },
        "_originalFormatting": Value::Object(original_formatting)
    }));
    if !array(field(Some(table), "propertyChanges")).is_empty() {
        attrs.insert(
            "tblPrChange".to_owned(),
            field(Some(table), "propertyChanges").unwrap().clone(),
        );
    }
    let row_spans = calculate_row_spans(table);
    let column_count = table_column_count(table);
    let rows = array(field(Some(table), "rows"))
        .iter()
        .enumerate()
        .map(|(row_index, row)| {
            project_row(
                row,
                table,
                row_index,
                &row_spans,
                &TableStyleContext {
                    column_count,
                    style: table_style.or(default_style),
                    borders,
                    margins: default_margins.as_ref(),
                    theme,
                },
            )
        })
        .collect();
    ProjectedTable { attrs, rows }
}

fn table_cell_story_id(parent: &str, table: usize, row: usize, cell: usize) -> String {
    format!("{parent}:t{table}:r{row}c{cell}")
}

/// Hands each block of a story the identity `visit_story` seeds it with.
#[derive(Clone, Copy, Default)]
struct BlockCursor {
    paragraph: usize,
    table: usize,
    sdt: usize,
}

impl BlockCursor {
    fn take(&mut self, story_id: &str, block: &Value) -> Option<String> {
        match string(field(Some(block), "type")).unwrap_or_default() {
            "rawXml" => None,
            "paragraph" => {
                let id = string(field(Some(block), "paraId"))
                    .filter(|value| !value.is_empty())
                    .map_or_else(|| format!("{story_id}:p{}", self.paragraph), str::to_owned);
                self.paragraph += 1;
                Some(id)
            }
            "table" => {
                let id = format!("{story_id}:t{}", self.table);
                self.table += 1;
                Some(id)
            }
            _ => {
                let id = format!("{story_id}:sdt{}", self.sdt);
                self.sdt += 1;
                Some(id)
            }
        }
    }
}

/// How many story blocks a suppressed field's cached result duplicates.
fn cached_result_block_count(data: &Value) -> Option<usize> {
    let blocks = array(field(field(Some(data), "structuredResult"), "blocks"));
    blocks
        .last()
        .is_some_and(|block| {
            string(field(Some(block), "type")) == Some("paragraph")
                && array(field(Some(block), "content")).iter().all(|node| {
                    string(field(Some(node), "type")) == Some("run")
                        && array(field(Some(node), "content")).iter().all(|entry| {
                            string(field(Some(entry), "type")) == Some("fieldChar")
                                && string(field(Some(entry), "charType")) == Some("end")
                                && field(Some(entry), "continuationId").is_some()
                        })
                })
        })
        .then(|| blocks.len())
}

/// Binds each suppressed field to the story blocks its cached result duplicates.
fn bind_field_result_blocks(
    units: &mut [InlineUnit],
    story_id: &str,
    blocks: &[Value],
    owner: usize,
    after_owner: BlockCursor,
    table_ids: &mut BTreeMap<usize, String>,
) {
    for unit in units {
        let UnitContent::Embed { kind, payload } = &mut unit.content else {
            continue;
        };
        // `instruction` is the key the render side suppresses on.
        if kind.as_str() != "field"
            || !numeric_field_instruction(
                payload
                    .get("instruction")
                    .and_then(Value::as_str)
                    .unwrap_or_default(),
            )
        {
            continue;
        }
        let Some(count) = payload
            .get("fieldData")
            .and_then(Value::as_str)
            .and_then(|data| serde_json::from_str::<Value>(data).ok())
            .as_ref()
            .and_then(cached_result_block_count)
        else {
            continue;
        };
        let Some(duplicated) = blocks.get(owner + 1..owner + 1 + count) else {
            continue;
        };
        let mut cursor = after_owner;
        let mut ids = Vec::with_capacity(duplicated.len());
        for (offset, block) in duplicated.iter().enumerate() {
            let is_table = string(field(Some(block), "type")) == Some("table");
            let Some(id) = cursor.take(story_id, block) else {
                continue;
            };
            if is_table {
                table_ids.insert(owner + 1 + offset, id.clone());
            }
            ids.push(Value::String(id));
        }
        if ids.is_empty() {
            continue;
        }
        payload.insert("fieldResultBlocks".to_owned(), Value::Array(ids));
    }
}

fn is_comment_reference(unit: &InlineUnit) -> bool {
    matches!(&unit.content, UnitContent::Embed { payload, .. }
        if payload.get("modelKind").and_then(Value::as_str) == Some("commentReference"))
}

fn add_comment_coverage(plan: &mut StoryPlan) {
    let mut offset = 0u32;
    for (index, unit) in plan.units.iter().enumerate() {
        if matches!(&unit.content, UnitContent::Embed {kind, ..} if kind == "bookmark") {
            continue;
        }
        let width = match &unit.content {
            UnitContent::Text(text) => utf16_len(text),
            UnitContent::Embed { .. } => 1,
        };
        // A reference mark follows its own range's end, so another range holds
        // it only when that range goes on past it: coincident ends stay equal.
        let next = is_comment_reference(unit).then(|| {
            plan.units[index + 1..].iter().find(|next| {
                !is_comment_reference(next)
                    && !matches!(&next.content, UnitContent::Embed {kind, ..} if kind == "bookmark")
            })
        });
        for comment_id in &unit.comments {
            if let Some(next) = next
                && !next.is_some_and(|next| next.comments.contains(comment_id))
            {
                continue;
            }
            let index = plan
                .comment_coverage
                .iter()
                .position(|(id, _)| id == comment_id);
            if let Some(index) = index {
                let ranges = &mut plan.comment_coverage[index].1;
                if let Some(previous) = ranges.last_mut()
                    && previous.1 == offset
                {
                    previous.1 = offset + width;
                } else {
                    ranges.push((offset, offset + width));
                }
            } else {
                plan.comment_coverage
                    .push((comment_id.clone(), vec![(offset, offset + width)]));
            }
        }
        offset += width;
    }
}

fn visit_story(
    context: &mut LoweringContext,
    story_id: String,
    source_blocks: &[Value],
    options: StoryOptions,
) {
    let plan_index = context.plans.len();
    context.plans.push(StoryPlan {
        story_id: story_id.clone(),
        units: Vec::new(),
        comment_coverage: Vec::new(),
    });
    let empty_story;
    let blocks = if source_blocks.is_empty() {
        empty_story = [json!({ "type": "paragraph", "content": [] })];
        &empty_story[..]
    } else {
        source_blocks
    };
    let mut cursor = BlockCursor::default();
    let mut result_table_ids = BTreeMap::new();
    let mut last_kind = None;
    for (block_index, block) in blocks.iter().enumerate() {
        let position = cursor;
        let Some(block_id) = cursor.take(&story_id, block) else {
            continue;
        };
        match string(field(Some(block), "type")).unwrap_or_default() {
            "paragraph" => {
                let (leading_breaks, trailing_breaks) = paragraph_flow_breaks(block);
                // Comment markers ahead of a leading break take it into their range.
                let content = array(field(Some(block), "content"));
                let mut marked = 0;
                for flow in leading_breaks {
                    for marker in &content[marked..flow.item] {
                        context.open_comments.mark(marker, plan_index);
                    }
                    marked = flow.item;
                    context.push(plan_index, flow.unit());
                }
                let (mut units, mut ppr) = paragraph_units(
                    block,
                    &context.styles,
                    None,
                    &context.source_json,
                    &mut context.open_comments,
                    plan_index,
                );
                ppr.insert("paraId".to_owned(), Value::String(block_id));
                bind_field_result_blocks(
                    &mut units,
                    &story_id,
                    blocks,
                    block_index,
                    cursor,
                    &mut result_table_ids,
                );
                context.plans[plan_index].units.extend(units);
                context.push(plan_index, embed_unit("pilcrow", ppr, &[], 1));
                for flow in trailing_breaks {
                    context.push(plan_index, flow.unit());
                }
                last_kind = Some("paragraph");
            }
            "table" => {
                let current_table = position.table;
                let table = project_table(
                    block,
                    &context.styles,
                    context.theme.as_ref(),
                    context.compatibility_mode,
                );
                let rows: Vec<Value> = table
                    .rows
                    .iter()
                    .enumerate()
                    .map(|(row_index, row)| {
                        json!({
                            "trPr": value_from_map(&row.attrs),
                            "cells": row.cells.iter().enumerate().map(|(cell_index, cell)| {
                                json!({
                                    "tcPr": value_from_map(&cell.attrs),
                                    "story": table_cell_story_id(
                                        &story_id,
                                        current_table,
                                        row_index,
                                        cell_index,
                                    )
                                })
                            }).collect::<Vec<_>>()
                        })
                    })
                    .collect();
                let tbl_pr = structural_attrs(table.attrs.clone(), &["columnWidths"]);
                let grid = table
                    .attrs
                    .get("columnWidths")
                    .and_then(Value::as_array)
                    .cloned()
                    .unwrap_or_default()
                    .into_iter()
                    .map(drop_nulls)
                    .collect::<Vec<_>>();
                let mut payload = map_from_value(json!({
                    "tblPr": value_from_map(&tbl_pr),
                    "grid": grid,
                    "rows": rows
                }));
                if let Some(id) = result_table_ids.remove(&block_index) {
                    payload.insert("blockId".to_owned(), Value::String(id));
                }
                context.push(plan_index, embed_unit("table", payload, &[], 1));
                let previous_table_formatting = context.styles.table_paragraph_formatting.take();
                for (row_index, row) in table.rows.into_iter().enumerate() {
                    for (cell_index, cell) in row.cells.into_iter().enumerate() {
                        context.styles.table_paragraph_formatting = cell.paragraph_formatting;
                        visit_story(
                            context,
                            table_cell_story_id(&story_id, current_table, row_index, cell_index),
                            &cell.content,
                            StoryOptions {
                                append_body_tail: false,
                            },
                        );
                    }
                }
                context.styles.table_paragraph_formatting = previous_table_formatting;
                last_kind = Some("table");
            }
            _ => {
                let child_story = block_id;
                let mut properties = sdt_properties_attrs(
                    field(Some(block), "properties").unwrap_or(&Value::Null),
                    &context.source_json,
                );
                properties.insert("story".to_owned(), Value::String(child_story.clone()));
                context.push(plan_index, embed_unit("blockSdt", properties, &[], 1));
                visit_story(
                    context,
                    child_story,
                    array(field(Some(block), "content")),
                    StoryOptions {
                        append_body_tail: false,
                    },
                );
                last_kind = Some("blockSdt");
            }
        }
    }
    if options.append_body_tail && matches!(last_kind, Some("table" | "blockSdt")) {
        context.push(
            plan_index,
            embed_unit(
                "pilcrow",
                map_from_value(json!({
                    "hangingIndent": false,
                    "paraId": format!("{story_id}:p{}", cursor.paragraph)
                })),
                &[],
                1,
            ),
        );
    }
    add_comment_coverage(&mut context.plans[plan_index]);
}

fn collect_font_entry(key: &str, value: &Value, fonts: &mut BTreeSet<String>) {
    if matches!(
        key,
        "fontFamily" | "listMarkerFontFamily" | "markerFontFamily"
    ) {
        match value {
            Value::String(name) if !name.trim().is_empty() => {
                fonts.insert(name.trim().to_owned());
            }
            Value::Object(slots) => {
                for key in ["ascii", "hAnsi", "eastAsia", "cs"] {
                    if let Some(name) = slots
                        .get(key)
                        .and_then(Value::as_str)
                        .filter(|name| !name.trim().is_empty())
                    {
                        fonts.insert(name.trim().to_owned());
                    }
                }
            }
            _ => {}
        }
    }
}

fn collect_fonts_from_value(value: &Value, fonts: &mut BTreeSet<String>) {
    match value {
        Value::Array(values) => {
            for value in values {
                collect_fonts_from_value(value, fonts);
            }
        }
        Value::Object(values) => {
            for (key, value) in values {
                collect_font_entry(key, value, fonts);
                collect_fonts_from_value(value, fonts);
            }
        }
        _ => {}
    }
}

fn collect_font_table_fonts(envelope: &docx_parse::S9WireEnvelope, fonts: &mut BTreeSet<String>) {
    for font in &envelope.document.package.font_table.fonts {
        if !font.name.trim().is_empty() {
            fonts.insert(font.name.trim().to_owned());
        }
        if let Some(name) = font
            .alt_name
            .as_deref()
            .filter(|name| !name.trim().is_empty())
        {
            fonts.insert(name.trim().to_owned());
        }
    }
}

fn units_to_raw_ops(
    units: Vec<InlineUnit>,
    referenced_fonts: &mut BTreeSet<String>,
) -> Result<Vec<RawOp>, String> {
    let mut ops = vec![RawOp::Delete { index: 0, len: 1 }];
    let mut index = 0u32;
    let mut text = String::new();
    let mut attrs = JsonObject::new();
    let flush = |ops: &mut Vec<RawOp>,
                 index: &mut u32,
                 text: &mut String,
                 attrs: &mut JsonObject|
     -> Result<(), String> {
        if text.is_empty() {
            return Ok(());
        }
        let inserted = std::mem::take(text);
        let len = utf16_len(&inserted);
        ops.push(RawOp::Insert {
            index: *index,
            text: inserted,
            attrs: yrs_attrs(std::mem::take(attrs))?,
        });
        *index += len;
        Ok(())
    };
    for unit in units {
        for (key, value) in &unit.attrs {
            collect_font_entry(key, value, referenced_fonts);
            collect_fonts_from_value(value, referenced_fonts);
        }
        if let UnitContent::Embed { payload, .. } = &unit.content {
            for (key, value) in payload {
                collect_font_entry(key, value, referenced_fonts);
                collect_fonts_from_value(value, referenced_fonts);
            }
        }
        match unit.content {
            UnitContent::Text(value) => {
                if text.is_empty() {
                    attrs = unit.attrs;
                } else if unit.attrs != attrs {
                    flush(&mut ops, &mut index, &mut text, &mut attrs)?;
                    attrs = unit.attrs;
                }
                text.push_str(&value);
            }
            UnitContent::Embed {
                kind,
                payload: values,
            } => {
                flush(&mut ops, &mut index, &mut text, &mut attrs)?;
                ops.push(RawOp::InsertEmbed {
                    index,
                    kind,
                    payload: payload(values)?,
                    attrs: yrs_attrs(unit.attrs)?,
                });
                index += 1;
            }
        }
    }
    flush(&mut ops, &mut index, &mut text, &mut attrs)?;
    Ok(ops)
}

/// Image embeds name a media part of the fingerprinted source as `media:<part>`
/// instead of carrying its bytes; lowering and the baseline resolve the part.
pub use docx_parse::media::MEDIA_REF_PREFIX;

/// Display data URL of every media part (a TIFF shows as PNG), by part path.
pub fn package_media(envelope: &docx_parse::S9WireEnvelope) -> HashMap<String, String> {
    envelope
        .document
        .package
        .media_entries
        .iter()
        .map(|(_, file)| (file.path.clone(), file.data_url.clone()))
        .collect()
}

/// Points source images at their media part. Parts with equal bytes share a
/// display URL; the first in package order wins, and the serializer keeps an
/// image's relationship to any part with the same bytes.
fn reference_media(units: &mut [InlineUnit], media: &HashMap<&str, &str>) {
    for unit in units {
        if let UnitContent::Embed { kind, payload } = &mut unit.content
            && kind == "image"
            && let Some(path) = payload
                .get("src")
                .and_then(Value::as_str)
                .and_then(|src| media.get(src).copied())
        {
            payload.insert(
                "src".to_owned(),
                Value::String(format!("{MEDIA_REF_PREFIX}{path}")),
            );
        }
    }
}

fn seed_plan(
    plan: StoryPlan,
    media: &HashMap<&str, &str>,
) -> Result<(String, Vec<RawOp>, BTreeSet<String>), String> {
    let StoryPlan {
        story_id,
        mut units,
        comment_coverage,
    } = plan;
    reference_media(&mut units, media);
    let bookmarks = take_bookmarks(&mut units)?;
    let mut referenced_fonts = BTreeSet::new();
    let mut ops = units_to_raw_ops(units, &mut referenced_fonts)?;
    ops.extend(bookmarks);
    if !comment_coverage.is_empty() {
        ops.extend(
            comment_coverage
                .into_iter()
                .map(|(id, ranges)| RawOp::SetComment {
                    id,
                    ranges,
                    author: String::new(),
                    date: String::new(),
                    body: Any::Null,
                }),
        );
    }
    Ok((story_id, ops, referenced_fonts))
}

/// Convert the parser's paragraph offsets to CRDT anchors only after all
/// text is inserted. Bookmarks inside an opaque control stay in its payload.
fn take_bookmarks(units: &mut Vec<InlineUnit>) -> Result<Vec<RawOp>, String> {
    let mut found = Vec::<(u32, JsonObject)>::new();
    let mut at = 0;
    let mut pm = 0;
    let mut spans = Vec::<(u32, u32, u32, u32)>::new();
    let mut leading = Vec::new();
    let mut head = true;
    for unit in units.iter_mut() {
        let width = match &unit.content {
            UnitContent::Text(text) => utf16_len(text),
            _ => 1,
        };
        if let UnitContent::Embed { kind, payload } = &mut unit.content {
            if kind == "bookmark" {
                let mut data = payload.clone();
                if matches!(string(data.get("kind")), Some("fieldseparate" | "fieldend")) {
                    data.insert("order".into(), json!(found.len()));
                    if !unit.attrs.is_empty() {
                        data.insert("attributes".into(), value_from_map(&unit.attrs));
                    }
                }
                found.push((at, data));
                continue;
            }
            if kind == "pilcrow" {
                if let Some(Value::Array(bookmarks)) = payload.remove("bookmarks") {
                    for bookmark in bookmarks {
                        let mut data = map_from_value(bookmark);
                        let offset = data
                            .remove("offset")
                            .and_then(|value| value.as_u64())
                            .unwrap_or(0) as u32;
                        let after = data
                            .remove("breaksAfter")
                            .and_then(|value| value.as_u64())
                            .unwrap_or(0) as usize;
                        let position = if offset == 0 && after > 0 {
                            *leading
                                .get(
                                    leading
                                        .len()
                                        .checked_sub(after)
                                        .ok_or("bookmark break position does not resolve")?,
                                )
                                .ok_or("bookmark break position does not resolve")?
                        } else {
                            spans
                                .iter()
                                .find(|(_, _, start, len)| offset < start + len)
                                .map_or(at, |(raw, width, start, _)| {
                                    raw + offset.saturating_sub(*start).min(*width)
                                })
                        };
                        found.push((position, data));
                    }
                }
                at += 1;
                pm = 0;
                spans.clear();
                leading.clear();
                head = true;
                continue;
            }
            if head && crate::segments::is_block_embed(kind) {
                if kind.ends_with("Break") {
                    leading.push(at);
                }
                at += 1;
                continue;
            }
        }
        if width > 0 {
            head = false;
            spans.push((at, width, pm, unit.pm_size));
            at += width;
            pm += unit.pm_size;
        }
    }
    units.retain(
        |unit| !matches!(&unit.content, UnitContent::Embed {kind, ..} if kind == "bookmark"),
    );
    found.sort_by_key(|(at, data)| (*at, data["id"].as_i64().unwrap_or(0), data["kind"] == "end"));
    found
        .into_iter()
        .map(|(index, data)| {
            Ok(RawOp::SetBookmark {
                index,
                data: any_from_value(value_from_map(&data))?,
            })
        })
        .collect()
}

fn entry_parts(entry: &Value) -> Option<(&str, &Value)> {
    let entry = entry.as_array()?;
    Some((entry.first()?.as_str()?, entry.get(1)?))
}

pub fn parse_docx_for_edit(bytes: &[u8]) -> Result<docx_parse::S9WireEnvelope, String> {
    docx_parse::parse_docx_s9_wire(bytes, docx_parse::S9ParseOptions::default())
        .map_err(|error| error.to_string())
}

#[cfg(feature = "wasm")]
pub(crate) fn referenced_fonts(
    envelope: &docx_parse::S9WireEnvelope,
) -> Result<Vec<String>, String> {
    let mut fonts = BTreeSet::new();
    collect_font_table_fonts(envelope, &mut fonts);
    let parsed = serde_json::to_value(&envelope.document).map_err(|error| error.to_string())?;
    collect_fonts_from_value(&parsed, &mut fonts);
    Ok(fonts.into_iter().collect())
}

/// What reseeding a field needs from its source package: the style lookup
/// the seed reads it with, and the theme and relationships its markup parses with.
pub struct PackageContext {
    styles: StyleResolver,
    pub(crate) theme: docx_parse::Theme,
    relationships: docx_parse::RelationshipMap,
    stories: docx_parse::StoryRelationships,
}

impl PackageContext {
    pub fn new(envelope: &docx_parse::S9WireEnvelope) -> Self {
        let package = &envelope.document.package;
        let styles = serde_json::to_value(&package.styles).ok();
        Self {
            styles: StyleResolver::new(styles.as_ref()),
            theme: package.theme.clone(),
            relationships: package.relationship_entries.iter().cloned().collect(),
            stories: envelope.story_relationships.clone(),
        }
    }

    /// The relationships markup in story `story_id` resolves with: its
    /// header, footer or notes part's own, or the document's.
    pub(crate) fn relationships(&self, story_id: &str) -> &docx_parse::RelationshipMap {
        let mut parts = story_id.split(':');
        let own = match (parts.next(), parts.next()) {
            (Some("hf"), Some(id)) => self
                .stories
                .header_footers
                .iter()
                .find(|(key, _)| key == id)
                .map(|(_, relationships)| relationships),
            (Some("fn"), _) => self.stories.footnotes.as_ref(),
            (Some("en"), _) => self.stories.endnotes.as_ref(),
            _ => None,
        };
        own.unwrap_or(&self.relationships)
    }
}

/// Every seed is written under this client, so a package seeds to the same
/// bytes and object ids on every engine instance (and costs one byte per id).
pub const SEED_CLIENT_ID: u64 = 0;

/// Seeds `document` from a parsed package under [`SEED_CLIENT_ID`].
pub fn seed_parsed_docx(
    document: &EditingDoc,
    envelope: docx_parse::S9WireEnvelope,
) -> Result<Vec<String>, String> {
    use yrs::Transact;
    use yrs::updates::decoder::Decode;
    document.set_package(Some(PackageContext::new(&envelope)));
    let seed = EditingDoc::new(SEED_CLIENT_ID);
    let fonts = seed_parsed_docx_in_place(&seed, envelope)?;
    let update = yrs::Update::decode_v1(&seed.encode_state_as_update_v1())
        .map_err(|error| error.to_string())?;
    document
        .yrs_doc()
        .transact_mut_with(document.client_id())
        .apply_update(update)
        .map_err(|error| error.to_string())?;
    Ok(fonts)
}

/// Seeds under `document`'s own client, for a view that never stores its
/// state; keeps update decoding out of the viewer build.
pub fn seed_parsed_docx_in_place(
    document: &EditingDoc,
    mut envelope: docx_parse::S9WireEnvelope,
) -> Result<Vec<String>, String> {
    let media = std::mem::take(&mut envelope.document.package.media_entries);
    let mut media_parts = HashMap::new();
    for (_, file) in &media {
        media_parts
            .entry(file.data_url.as_str())
            .or_insert(file.path.as_str());
    }
    let mut referenced_fonts = BTreeSet::new();
    collect_font_table_fonts(&envelope, &mut referenced_fonts);
    let parsed = serde_json::to_value(&envelope.document).map_err(|error| error.to_string())?;
    collect_fonts_from_value(&parsed, &mut referenced_fonts);
    let source_json = if needs_source_json(&parsed) {
        let serialized =
            serde_json::to_string(&envelope.document).map_err(|error| error.to_string())?;
        let ordered: OrderedValue =
            serde_json::from_str(&serialized).map_err(|error| error.to_string())?;
        let mut values = BTreeMap::new();
        ordered.collect_source_json(&mut values);
        values
    } else {
        BTreeMap::new()
    };
    drop(envelope);
    let package =
        field(Some(&parsed), "package").ok_or_else(|| "parsed DOCX has no package".to_owned())?;
    let compatibility_mode = compatibility_mode_from_package(Some(package));
    let mut context = LoweringContext {
        styles: StyleResolver::new(field(Some(package), "styles")),
        theme: field(Some(package), "theme").cloned(),
        source_json: Arc::new(source_json),
        plans: Vec::new(),
        compatibility_mode,
        open_comments: OpenComments::default(),
    };
    visit_story(
        &mut context,
        "body".to_owned(),
        array(field(field(Some(package), "document"), "content")),
        StoryOptions {
            append_body_tail: true,
        },
    );
    for entry in array(field(Some(package), "headerEntries")) {
        let Some((relationship_id, part)) = entry_parts(entry) else {
            continue;
        };
        visit_story(
            &mut context,
            format!("hf:{relationship_id}"),
            array(field(Some(part), "content")),
            StoryOptions {
                append_body_tail: false,
            },
        );
    }
    for entry in array(field(Some(package), "footerEntries")) {
        let Some((relationship_id, part)) = entry_parts(entry) else {
            continue;
        };
        let story_id = format!("hf:{relationship_id}");
        if context.plans.iter().any(|plan| plan.story_id == story_id) {
            continue;
        }
        visit_story(
            &mut context,
            story_id,
            array(field(Some(part), "content")),
            StoryOptions {
                append_body_tail: false,
            },
        );
    }
    for (key, prefix) in [("footnotes", "fn"), ("endnotes", "en")] {
        for note in array(field(Some(package), key)) {
            let Some(id) = field(Some(note), "id") else {
                continue;
            };
            visit_story(
                &mut context,
                format!("{prefix}:{}", js_string(id)),
                array(field(Some(note), "content")),
                StoryOptions {
                    append_body_tail: false,
                },
            );
        }
    }
    let comment_ops = array(field(field(Some(package), "document"), "comments"))
        .iter()
        .filter_map(|comment| {
            let id = field(Some(comment), "id")?;
            Some((js_string(id), comment))
        })
        .map(|(id, comment)| {
            let mut fields = Vec::new();
            for key in ["author", "date", "parentId", "done"] {
                if let Some(value) = field(Some(comment), key) {
                    let value = if key == "parentId" && !value.is_null() {
                        Value::String(js_string(value))
                    } else {
                        value.clone()
                    };
                    fields.push((key.to_owned(), any_from_value(value)?));
                }
            }
            if let Some(content) = field(Some(comment), "content") {
                fields.push(("body".to_owned(), any_from_value(content.clone())?));
            }
            Ok(RawOp::PatchComment { id, fields })
        })
        .collect::<Result<Vec<_>, String>>()?;
    drop(parsed);
    document
        .create_empty_stories(
            &context
                .plans
                .iter()
                .map(|plan| plan.story_id.clone())
                .collect::<Vec<_>>(),
        )
        .map_err(|error| error.to_string())?;
    let mut batches = Vec::with_capacity(context.plans.len());
    for plan in context.plans {
        let (story_id, ops, fonts) = seed_plan(plan, &media_parts)?;
        batches.push((story_id, ops));
        referenced_fonts.extend(fonts);
    }
    batches.push(("body".to_owned(), comment_ops));
    document
        .apply_raw_story_batches(batches, &EditCtx::local(String::new(), String::new()))
        .map_err(|error| error.to_string())?;
    Ok(referenced_fonts.into_iter().collect())
}

pub fn seed_from_docx(document: &EditingDoc, bytes: &[u8]) -> Result<(), String> {
    let envelope = parse_docx_for_edit(bytes)?;
    seed_parsed_docx(document, envelope).map(|_| ())
}

#[cfg(test)]
mod tests {
    fn seed_body(blocks: &[Value]) -> EditingDoc {
        let mut context = LoweringContext {
            styles: StyleResolver::new(None),
            theme: None,
            source_json: Arc::new(BTreeMap::new()),
            plans: Vec::new(),
            compatibility_mode: 12,
            open_comments: OpenComments::default(),
        };
        visit_story(
            &mut context,
            "body".to_owned(),
            blocks,
            StoryOptions {
                append_body_tail: false,
            },
        );
        let document = EditingDoc::new(74101);
        document
            .create_empty_stories(
                &context
                    .plans
                    .iter()
                    .map(|plan| plan.story_id.clone())
                    .collect::<Vec<_>>(),
            )
            .unwrap();
        let batches = context
            .plans
            .into_iter()
            .map(|plan| {
                let (story_id, ops, _) = seed_plan(plan, &HashMap::new()).unwrap();
                (story_id, ops)
            })
            .collect();
        document
            .apply_raw_story_batches(batches, &EditCtx::local(String::new(), String::new()))
            .unwrap();
        document
    }

    fn rendered(document: &EditingDoc) -> (usize, String) {
        let blocks = crate::bridge::yrs_doc_to_layout_blocks(
            document,
            "body",
            &crate::bridge::RenderEnv::default(),
        )
        .unwrap();
        let json = serde_json::to_string(&blocks).unwrap();
        (blocks.len(), json)
    }

    fn run(text: &str) -> Value {
        json!({"type":"run","content":[{"type":"text","text":text}]})
    }

    fn block_field(instruction: &str, blocks: &[Value]) -> Value {
        json!({
            "type":"complexField", "fieldType":"UNKNOWN", "instruction":instruction,
            "fieldCode":[], "fieldResult":[run("Cached first")],
            "structuredResult":{
                "inline":[{"type":"hyperlink","anchor":"bookmark","children":[run("Cached first")]}],
                "blocks":blocks
            }
        })
    }

    #[test]
    fn numeric_fields_hide_cached_paragraphs_without_changing_the_story() {
        // `w14:paraId` is optional and plenty of real documents carry none, so
        // the binding cannot key on it.
        for (instruction, para_ids) in [("0", true), ("TOC", true), ("0", false), ("TOC", false)] {
            let id = |name: &str| if para_ids { json!(name) } else { json!(null) };
            let cached =
                json!({"type":"paragraph","paraId":id("cached"),"content":[run("Cached second") ]});
            let end = json!({"type":"paragraph","paraId":id("end"),"content":[]});
            let field = block_field(instruction, &[cached.clone(), end.clone()]);
            let document = seed_body(&[
                json!({"type":"paragraph","paraId":id("owner"),"content":[field]}),
                cached,
                end,
                json!({"type":"paragraph","paraId":id("after"),"content":[run("After")]}),
            ]);
            let before = crate::story_checksum(&document, "body").unwrap();
            let (count, output) = rendered(&document);
            let label = format!("{instruction} paraIds={para_ids}");
            assert_eq!(count, if instruction == "0" { 2 } else { 4 }, "{label}");
            assert_eq!(output.contains("Cached"), instruction != "0", "{label}");
            assert!(output.contains("After"), "{label}");
            assert_eq!(
                before,
                crate::story_checksum(&document, "body").unwrap(),
                "{label}"
            );
        }
    }

    #[test]
    fn numeric_fields_hide_non_paragraph_cached_blocks() {
        for instruction in ["0", "TOC"] {
            let cell = json!({"type":"tableCell","content":[
                json!({"type":"paragraph","content":[run("Cached cell")]})
            ]});
            let table = json!({"type":"table","rows":[{"type":"tableRow","cells":[cell]}]});
            let sdt = json!({"type":"blockSdt","properties":{},"content":[
                json!({"type":"paragraph","content":[run("Cached sdt")]})
            ]});
            let end = json!({"type":"paragraph","content":[]});
            let field = block_field(instruction, &[table.clone(), sdt.clone(), end.clone()]);
            let document = seed_body(&[
                json!({"type":"paragraph","content":[field]}),
                table,
                sdt,
                end,
                json!({"type":"paragraph","content":[run("After")]}),
            ]);
            let before = crate::story_checksum(&document, "body").unwrap();
            let (count, output) = rendered(&document);
            assert_eq!(
                count,
                if instruction == "0" { 2 } else { 5 },
                "{instruction}"
            );
            assert_eq!(
                output.contains("Cached"),
                instruction != "0",
                "{instruction}"
            );
            assert!(output.contains("After"), "{instruction}");
            assert_eq!(
                before,
                crate::story_checksum(&document, "body").unwrap(),
                "{instruction}"
            );
        }
    }

    #[test]
    fn splitting_a_cached_paragraph_shows_the_authored_half_only() {
        let cached = json!({"type":"paragraph","content":[run("Cached second")]});
        let end = json!({"type":"paragraph","content":[]});
        let document = seed_body(&[
            json!({"type":"paragraph","content":[block_field("0", &[cached.clone(), end.clone()])]}),
            cached,
            end,
            json!({"type":"paragraph","content":[run("After")]}),
        ]);
        assert_eq!(rendered(&document).0, 2);
        // The field embed, the owner's pilcrow, then the cached text: pressing
        // Enter at the end of the cached paragraph and typing into the new one.
        let cached_pilcrow = 2 + "Cached second".encode_utf16().count() as u32;
        let ctx = EditCtx::local("Ada".to_owned(), "2026-01-01T00:00:00Z".to_owned());
        document
            .split_paragraph(&ctx, crate::Position::new("body", cached_pilcrow), None)
            .unwrap();
        document
            .insert_text(
                &ctx,
                crate::Position::new("body", cached_pilcrow + 1),
                "Authored",
                crate::FormatPolicy::Plain,
            )
            .unwrap();
        let (count, output) = rendered(&document);
        assert_eq!(count, 3);
        assert!(output.contains("Authored"));
        assert!(output.contains("After"));
        assert!(!output.contains("Cached"));
    }

    #[test]
    fn numeric_field_detection_preserves_formulas_and_named_fields() {
        assert!(numeric_field_instruction(" 123 "));
        for instruction in ["", "= 0", "QUOTE 0", "PAGE", "CustomField", "123abc"] {
            assert!(!numeric_field_instruction(instruction));
        }
    }

    #[test]
    fn complex_field_results_keep_hyperlink_units_and_style() {
        let link = json!({"type":"hyperlink","anchor":"_Toc1","children":[{"type":"run","formatting":{"styleId":"Hyperlink"},"content":[{"type":"text","text":"Heading"}]}]});
        let value = json!({"type":"complexField","fieldType":"TOC","instruction":"TOC", "fieldCode":[], "fieldResult":link["children"], "structuredResult":{"inline":[link, {"type":"simpleField","fieldType":"PAGE","instruction":" PAGE ","content":[{"type":"run","content":[{"type":"text","text":"1"}]}]}]}});
        let styles = StyleResolver::new(Some(
            &json!({"styles":[{"type":"character","styleId":"Hyperlink","rPr":{"color":{"rgb":"0563C1"}}}]}),
        ));
        let (units, _) = paragraph_units(
            &json!({"content":[value]}),
            &styles,
            None,
            &BTreeMap::new(),
            &mut OpenComments::default(),
            0,
        );
        assert!(matches!(&units[0].content, UnitContent::Text(text) if text == "Heading"));
        assert_eq!(units[0].attrs["hyperlink"]["href"], json!("#_Toc1"));
        assert_eq!(units[0].attrs["textColor"]["rgb"], json!("0563C1"));
        let UnitContent::Embed {
            payload: nested, ..
        } = &units[1].content
        else {
            panic!("missing nested field")
        };
        assert_eq!(nested["fieldType"], json!("PAGE"));
        let UnitContent::Embed { payload, .. } = &units[2].content else {
            panic!("missing field")
        };
        assert_eq!(payload["displayText"], json!(""));
        assert!(
            payload["fieldData"]
                .as_str()
                .unwrap()
                .contains("structuredResult")
        );
    }

    #[test]
    fn comprehensive_native_layout_accepts_authored_page_number_start() {
        let bytes = include_bytes!(
            "../../betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx"
        );
        let parsed =
            docx_parse::parse_docx_s9_wire(bytes, docx_parse::S9ParseOptions::default()).unwrap();
        let package = parsed.document.package;
        let mut sections: Vec<Value> = package
            .document
            .sections
            .unwrap()
            .into_iter()
            .map(|section| json!({"properties":section.properties}))
            .collect();
        assert!(
            sections.iter().any(
                |section| section["properties"]["pageNumbering"]["start"].as_f64() == Some(1.0)
            )
        );
        sections.push(json!({"properties":package.document.final_section_properties}));
        let request = json!({"bodyStory":"body", "options":{"pageGap":24}, "regions":{"sections":sections, "settings":package.settings}, "renderEnv":{}}).to_string();
        let engine = crate::EngineSession::new(74003);
        seed_from_docx(engine.doc(), bytes).unwrap();
        engine.layout_font_requirements_json(&request).unwrap();
        let layout: Value =
            serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                .unwrap();
        assert!(!layout["layout"]["pages"].as_array().unwrap().is_empty());
    }

    #[test]
    fn raw_blocks_seed_no_content_control_or_child_story() {
        let mut context = LoweringContext {
            styles: StyleResolver::new(None),
            theme: None,
            source_json: Arc::new(BTreeMap::new()),
            plans: Vec::new(),
            compatibility_mode: 12,
            open_comments: OpenComments::default(),
        };
        visit_story(
            &mut context,
            "body".to_owned(),
            &[
                json!({"type":"paragraph","content":[]}),
                json!({"type":"rawXml","xml":"<x:block/>"}),
                json!({"type":"paragraph","content":[]}),
            ],
            StoryOptions {
                append_body_tail: true,
            },
        );
        assert_eq!(context.plans.len(), 1);
        assert_eq!(context.plans[0].units.len(), 2);
        assert!(context.plans[0].units.iter().all(
            |unit| matches!(&unit.content, UnitContent::Embed { kind, .. } if kind == "pilcrow")
        ));
    }

    #[test]
    fn run_boundaries_are_kept_only_for_what_merging_would_lose() {
        let boundaries = |content: Value| {
            paragraph_units(
                &json!({ "content": content }),
                &StyleResolver::new(None),
                None,
                &BTreeMap::new(),
                &mut OpenComments::default(),
                0,
            )
            .1
            .get("_originalRunBoundaries")
            .cloned()
        };
        let plain =
            json!({"type":"run","formatting":{"bold":true},"content":[{"type":"text","text":"A"}]});
        assert_eq!(
            boundaries(json!([plain, {"type":"rawXml","xml":"<x:mark/>"}, plain])),
            None
        );
        let tracked = json!({"type":"run","propertyChanges":[{"id":1}],"content":[{"type":"text","text":"B"}]});
        let empty = json!({"type":"run","formatting":{"bold":true},"content":[]});
        let kept = boundaries(json!([plain, tracked])).unwrap();
        assert_eq!(kept[0], json!({"text":"A","marksKey":"bold:{}"}));
        assert_eq!(kept[1]["propertyChanges"], json!([{"id":1}]));
        let kept = boundaries(json!([plain, empty])).unwrap();
        assert_eq!(kept[1], json!({"text":"","formatting":{"bold":true}}));
    }

    #[test]
    fn numbering_indents_precede_styles_and_ignore_zero_first_line() {
        let style_data = json!({"styles":[{"styleId":"List","type":"paragraph","pPr":{"indentLeft":720,"indentFirstLine":180,"hangingIndent":false}}]});
        for styles in [
            StyleResolver::new(Some(&style_data)),
            StyleResolver::new(None),
        ] {
            for direct in [
                json!({}),
                json!({"indentFirstLine":0}),
                json!({"indentFirstLine":0,"hangingIndent":true}),
            ] {
                let mut formatting = direct;
                formatting["styleId"] = json!("List");
                let properties = paragraph_attrs(
                    &json!({"formatting":formatting,"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
                    &styles,
                    &[],
                    &[],
                    None,
                );
                assert_eq!(properties["indentLeft"], json!(1440));
                assert_eq!(properties["indentFirstLine"], json!(-360));
                assert_eq!(properties["hangingIndent"], json!(true));
            }
            let properties = paragraph_attrs(
                &json!({"formatting":{"styleId":"List","indentLeft":0,"indentFirstLine":240,"hangingIndent":false},"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
                &styles,
                &[],
                &[],
                None,
            );
            assert_eq!(properties["indentLeft"], json!(0));
            assert_eq!(properties["indentFirstLine"], json!(240));
            assert_eq!(properties["hangingIndent"], json!(false));
        }
    }

    #[test]
    fn first_line_value_and_kind_share_one_source() {
        let hanging_styles = StyleResolver::new(Some(
            &json!({"styles":[{"styleId":"Normal","type":"paragraph","default":true,"pPr":{"indentLeft":1450,"indentFirstLine":-730,"hangingIndent":true}}]}),
        ));
        let first_styles = StyleResolver::new(Some(
            &json!({"styles":[{"styleId":"Normal","type":"paragraph","default":true,"pPr":{"indentLeft":1450,"indentFirstLine":720}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":2160,"indentFirstLine":720},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(2160));
        assert_eq!(properties["indentFirstLine"], json!(720));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":720,"indentFirstLine":0},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(0));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":1425},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(1425));
        assert_eq!(properties["indentFirstLine"], json!(-730));
        assert_eq!(properties["hangingIndent"], json!(true));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":2160,"indentFirstLine":-720,"hangingIndent":true},"content":[]}),
            &first_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(-720));
        assert_eq!(properties["hangingIndent"], json!(true));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentFirstLine":720},"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(720));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{},"listRendering":{"indentLeft":1440,"indentFirstLine":300},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(300));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentFirstLine":0},"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(-360));
        assert_eq!(properties["hangingIndent"], json!(true));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentFirstLine":-720,"hangingIndent":true},"listRendering":{"indentLeft":2145},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(2145));
        assert_eq!(properties["indentFirstLine"], json!(-720));
        assert_eq!(properties["hangingIndent"], json!(true));
    }

    #[test]
    fn derived_first_line_without_flag_clears_base_hanging() {
        let styles = StyleResolver::new(Some(
            &json!({"docDefaults":{"pPr":{"indentFirstLine":-730,"hangingIndent":true}},"styles":[{"styleId":"Derived","type":"paragraph","pPr":{"indentFirstLine":200}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"styleId":"Derived"},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(200));
        assert_eq!(properties["hangingIndent"], json!(false));
        let styles = StyleResolver::new(Some(
            &json!({"docDefaults":{"pPr":{"indentFirstLine":-730,"hangingIndent":true}},"styles":[{"styleId":"Derived","type":"paragraph","pPr":{"indentFirstLine":0}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"styleId":"Derived"},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(0));
        assert_eq!(properties["hangingIndent"], json!(false));
        let styles = StyleResolver::new(Some(
            &json!({"docDefaults":{"pPr":{"indentFirstLine":200}},"styles":[{"styleId":"Derived","type":"paragraph","pPr":{"indentFirstLine":-360,"hangingIndent":true}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"styleId":"Derived"},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(-360));
        assert_eq!(properties["hangingIndent"], json!(true));
    }

    #[test]
    fn indent_kind_pairs_merge_atomically_including_zero() {
        let base = json!({"indentFirstLine":-730,"hangingIndent":true});
        let derived = json!({"indentFirstLine":200});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(200));
        assert!(merged.get("hangingIndent").is_none());
        let derived = json!({"indentFirstLine":0});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(0));
        assert!(merged.get("hangingIndent").is_none());
        let base = json!({"indentFirstLine":200});
        let derived = json!({"hangingIndent":true});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(200));
        assert!(merged.get("hangingIndent").is_none());
        let base = json!({"indentFirstLine":-730,"hangingIndent":true});
        let derived = json!({"indentLeft":720});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(-730));
        assert_eq!(merged["hangingIndent"], json!(true));
    }

    #[test]
    fn parsed_indent_xml_seeds_matching_attrs() {
        fn ppr(xml: &str) -> Value {
            let limits = docx_parse::xml::ParseLimits::default();
            let mut budget = docx_parse::xml::ParseBudget::new(&limits);
            let root = docx_parse::xml::parse_xml(xml.as_bytes(), "formatting.xml", &mut budget)
                .unwrap()
                .root()
                .unwrap()
                .clone();
            serde_json::to_value(docx_parse::parse_paragraph_properties(Some(&root), None).unwrap())
                .unwrap()
        }
        let style_ppr = ppr(r#"<w:pPr><w:ind w:left="1450" w:hanging="730"/></w:pPr>"#);
        assert_eq!(style_ppr["indentFirstLine"], json!(-730.0));
        assert_eq!(style_ppr["hangingIndent"], json!(true));
        let direct_ppr = ppr(r#"<w:pPr><w:ind w:left="2160" w:firstLine="720"/></w:pPr>"#);
        assert_eq!(direct_ppr["indentFirstLine"], json!(720.0));
        assert!(direct_ppr.get("hangingIndent").is_none());
        let styles = StyleResolver::new(Some(
            &json!({"styles":[{"styleId":"Normal","type":"paragraph","default":true,"pPr":style_ppr}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting": direct_ppr, "content": []}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(2160.0));
        assert_eq!(properties["indentFirstLine"], json!(720.0));
        assert_eq!(properties["hangingIndent"], json!(false));
    }

    #[test]
    fn numbering_level_marker_format_preserves_explicit_off() {
        let styles = StyleResolver::new(None);
        let properties = paragraph_attrs(
            &json!({"formatting":{},"listRendering":{"marker":"1.","numFmt":"decimal","markerBold":false,"markerItalic":false,"markerColor":{"rgb":"000000"},"markerFontFamily":"Times New Roman","markerFontSize":12.0},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["listMarker"], json!("1."));
        assert_eq!(properties["listMarkerBold"], json!(false));
        assert_eq!(properties["listMarkerItalic"], json!(false));
        assert_eq!(properties["listMarkerColor"], json!({"rgb":"000000"}));
        assert_eq!(properties["listMarkerFontFamily"], json!("Times New Roman"));
        assert_eq!(properties["listMarkerFontSize"], json!(12.0));
    }

    use super::*;

    #[test]
    fn source_images_reference_their_part_and_resolve_at_lowering() {
        let src = "data:image/png;base64,AQID";
        for with_field in [false, true] {
            let mut envelope = parse_docx_for_edit(include_bytes!(
                "../../../apps/demo/public/betteroffice-demo.docx"
            ))
            .unwrap();
            let mut content = vec![json!({
                "type": "run",
                "formatting": {"fontFamily": {"ascii": "Image Caption"}},
                "content": [{"type": "drawing", "image": {
                    "type": "image", "rId": "rIdImage", "src": src,
                    "size": {"width": 914400, "height": 457200},
                    "wrap": {"type": "inline"}
                }}]
            })];
            if with_field {
                content.push(json!({
                    "type": "simpleField", "instruction": " PAGE ", "fieldType": "PAGE",
                    "content": [{"type": "run", "content": [{"type": "text", "text": "1"}]}]
                }));
            }
            envelope.document.package.document.content = serde_json::from_value(json!([{
                "type": "paragraph", "paraId": "image", "content": content
            }]))
            .unwrap();
            envelope.document.package.media_entries = vec![(
                "word/media/image.png".to_owned(),
                Arc::new(docx_parse::media::MediaFile {
                    path: "word/media/image.png".to_owned(),
                    filename: Some("image.png".to_owned()),
                    mime_type: "image/png".to_owned(),
                    base64: "AQID".to_owned(),
                    data_url: src.to_owned(),
                }),
            )];
            let media = package_media(&envelope);
            let engine = crate::EngineSession::new(7);
            let fonts = seed_parsed_docx(engine.doc(), envelope).unwrap();
            assert!(fonts.iter().any(|font| font == "Image Caption"));
            let state = engine.doc().encode_state_as_update_v1();
            let contains =
                |needle: &[u8]| state.windows(needle.len()).any(|window| window == needle);
            assert!(contains(b"media:word/media/image.png") && !contains(b"AQID"));
            let image_src = |engine: &crate::EngineSession| {
                engine
                    .with_lowered_story("body", &crate::bridge::RenderEnv::default(), |blocks| {
                        let docx_layout::types::LayoutBlock::Paragraph(paragraph) = &blocks[0]
                        else {
                            panic!("image paragraph must remain a paragraph");
                        };
                        paragraph.runs.iter().find_map(|run| match run {
                            docx_layout::types::Run::Image(image) => {
                                assert_eq!((image.width, image.height), (96.0, 48.0));
                                Some(image.src.clone())
                            }
                            _ => None,
                        })
                    })
                    .unwrap()
                    .unwrap()
            };
            // Without the package the reference stays unresolved and paints nothing.
            assert_eq!(image_src(&engine), "media:word/media/image.png");
            engine.set_media(media);
            assert_eq!(image_src(&engine), src);
        }
    }

    #[test]
    fn seeds_are_byte_identical_under_the_fixed_seed_client() {
        use yrs::{ReadTxn, Transact};
        let seeded = |client_id| {
            let document = EditingDoc::new(client_id);
            seed_from_docx(
                &document,
                include_bytes!("../../../apps/demo/public/betteroffice-demo.docx"),
            )
            .unwrap();
            document
        };
        let (left, right) = (seeded(11), seeded(22));
        assert_eq!(
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1()
        );
        let clients: Vec<u64> = left
            .yrs_doc()
            .transact()
            .state_vector()
            .iter()
            .map(|(client, _)| client.get())
            .collect();
        assert_eq!(clients, [SEED_CLIENT_ID]);
    }

    #[test]
    fn source_json_preserves_wire_order_with_js_number_formatting() {
        let ordered: OrderedValue =
            serde_json::from_str(r#"{"type":"shape","z":1.0,"nested":{"b":2,"a":3}}"#).unwrap();
        let value = ordered.value();
        let mut source = BTreeMap::new();
        ordered.collect_source_json(&mut source);

        assert_eq!(
            source_json(&value, &source),
            r#"{"type":"shape","z":1,"nested":{"b":2,"a":3}}"#
        );
    }

    fn widow_control_styles() -> Value {
        // Style chains are already merged: Body carries Normal's authored off,
        // while docDefaults sits under a style that leaves the toggle absent.
        json!({
            "docDefaults": { "pPr": { "widowControl": false } },
            "styles": [
                { "styleId": "Normal", "type": "paragraph", "default": true, "pPr": {} },
                { "styleId": "Body", "type": "paragraph", "pPr": { "widowControl": false } },
                { "styleId": "Quote", "type": "paragraph", "pPr": { "widowControl": true } }
            ]
        })
    }

    fn seeded_widow_control(styles: &StyleResolver, formatting: Value) -> Option<Value> {
        paragraph_attrs(
            &json!({ "formatting": formatting, "content": [] }),
            styles,
            &[],
            &[],
            None,
        )
        .get("widowControl")
        .cloned()
    }

    fn snap_grid_styles() -> Value {
        json!({
            "docDefaults": { "pPr": {} },
            "styles": [
                { "styleId": "Normal", "type": "paragraph", "default": true, "pPr": {} },
                { "styleId": "Body", "type": "paragraph", "pPr": { "snapToGrid": false } },
                { "styleId": "Quote", "type": "paragraph", "pPr": { "snapToGrid": true } }
            ]
        })
    }

    fn seeded_snap_to_grid(styles: &StyleResolver, formatting: Value) -> Option<Value> {
        paragraph_attrs(
            &json!({ "formatting": formatting, "content": [] }),
            styles,
            &[],
            &[],
            None,
        )
        .get("snapToGrid")
        .cloned()
    }

    #[test]
    fn snap_to_grid_is_seeded_from_the_style_and_direct_formatting() {
        let styles = StyleResolver::new(Some(&snap_grid_styles()));

        assert_eq!(seeded_snap_to_grid(&styles, json!({})), Some(Value::Null));
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Body" })),
            Some(Value::Bool(false))
        );
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Quote" })),
            Some(Value::Bool(true))
        );
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Body", "snapToGrid": true })),
            Some(Value::Bool(true)),
            "a direct on overrides a style that opts out"
        );
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Quote", "snapToGrid": false })),
            Some(Value::Bool(false)),
            "a direct off overrides a style that turns the toggle back on"
        );
    }

    #[test]
    fn widow_control_is_seeded_from_doc_defaults_the_style_and_direct_formatting() {
        let styles = StyleResolver::new(Some(&widow_control_styles()));

        assert_eq!(
            seeded_widow_control(&styles, json!({})),
            Some(Value::Bool(false)),
            "docDefaults reaches a paragraph whose style is silent"
        );
        assert_eq!(
            seeded_widow_control(&styles, json!({ "styleId": "Body" })),
            Some(Value::Bool(false))
        );
        assert_eq!(
            seeded_widow_control(&styles, json!({ "styleId": "Quote" })),
            Some(Value::Bool(true))
        );
        assert_eq!(
            seeded_widow_control(&styles, json!({ "styleId": "Body", "widowControl": true })),
            Some(Value::Bool(true))
        );
        assert_eq!(
            seeded_widow_control(
                &styles,
                json!({ "styleId": "Quote", "widowControl": false })
            ),
            Some(Value::Bool(false)),
            "a direct off overrides a style that turns the toggle back on"
        );
    }

    #[test]
    fn note_ref_marks_seed_no_story_unit_and_land_in_the_run_boundary() {
        let styles = StyleResolver::new(None);
        for (content_type, note_type) in [
            ("footnoteRefMark", "footnote"),
            ("endnoteRefMark", "endnote"),
        ] {
            let run = json!({
                "type": "run",
                "formatting": { "styleId": "FootnoteReference" },
                "content": [{ "type": content_type }, { "type": content_type }],
            });
            let units = run_to_units(&run, None, &styles, &[], &BTreeMap::new());
            assert!(units.is_empty());
            let boundary = run_boundary(&run, &units, &BTreeMap::new()).unwrap();
            assert_eq!(
                boundary.get("noteMarks"),
                Some(&json!([note_type, note_type]))
            );
            assert_eq!(boundary.get("text"), Some(&Value::String(String::new())));
        }
    }

    #[test]
    fn flow_breaks_land_in_the_run_boundary_at_their_text_offset() {
        let styles = StyleResolver::new(None);
        let wrapping = json!({
            "type": "run",
            "content": [
                { "type": "text", "text": "AB" },
                { "type": "break", "breakType": "page" },
                { "type": "break", "breakType": "textWrapping" },
            ],
        });
        let units = run_to_units(&wrapping, None, &styles, &[], &BTreeMap::new());
        // A wrapping break is an inline unit, so the run keeps no boundary.
        assert!(run_boundary(&wrapping, &units, &BTreeMap::new()).is_none());

        let run = json!({
            "type": "run",
            "content": [
                { "type": "text", "text": "AB" },
                { "type": "break", "breakType": "page" },
                { "type": "text", "text": "C" },
                { "type": "break", "breakType": "column" },
            ],
        });
        let units = run_to_units(&run, None, &styles, &[], &BTreeMap::new());
        assert!(run_boundary(&run, &units, &BTreeMap::new()).is_none());
        let text_units: Vec<_> = units
            .into_iter()
            .filter(|unit| matches!(unit.content, UnitContent::Text(_)))
            .collect();
        let boundary = run_boundary(&run, &text_units, &BTreeMap::new()).unwrap();
        assert_eq!(boundary.get("text"), Some(&json!("ABC")));
        assert_eq!(
            boundary.get("breaks"),
            Some(&json!([
                { "offset": 2, "type": "page" },
                { "offset": 3, "type": "column" },
            ]))
        );
    }

    #[test]
    fn breaks_before_a_paragraphs_text_lead_it_and_keep_their_tracked_change() {
        let describe = |breaks: Vec<FlowBreak>| {
            breaks
                .into_iter()
                .map(|flow| {
                    let unit = flow.unit();
                    let UnitContent::Embed { kind, payload } = unit.content else {
                        panic!("a break is an embed")
                    };
                    format!(
                        "{kind}{}{}",
                        if payload.contains_key("leading") {
                            "^"
                        } else {
                            ""
                        },
                        if unit.attrs.contains_key("del") {
                            "-"
                        } else {
                            ""
                        }
                    )
                })
                .collect::<Vec<_>>()
        };
        let page = json!({"type": "break", "breakType": "page"});
        let column = json!({"type": "break", "breakType": "column"});
        let paragraph = json!({"type": "paragraph", "content": [
            {"type": "deletion", "info": {"id": 7, "author": "A"}, "content": [
                {"type": "run", "content": [page]}
            ]},
            {"type": "run", "content": [column, {"type": "text", "text": "x"}, page]},
        ]});
        let (leading, trailing) = paragraph_flow_breaks(&paragraph);
        assert_eq!(describe(leading), ["pageBreak^-", "columnBreak^"]);
        assert_eq!(describe(trailing), ["pageBreak"]);

        // Without text the breaks follow the paragraph's mark.
        let (leading, trailing) = paragraph_flow_breaks(
            &json!({"type": "paragraph", "content": [{"type": "run", "content": [page]}]}),
        );
        assert!(leading.is_empty());
        assert_eq!(describe(trailing), ["pageBreak"]);
    }

    #[test]
    fn a_comment_reference_before_a_break_does_not_make_it_trailing() {
        let reference = json!({"type": "commentReference", "id": 1});
        let page = json!({"type": "break", "breakType": "page"});
        for content in [
            vec![
                reference.clone(),
                page.clone(),
                json!({"type": "text", "text": "Heading"}),
            ],
            vec![page, reference],
        ] {
            let (leading, trailing) = paragraph_flow_breaks(&json!({
                "type": "paragraph", "content": [{"type": "run", "content": content}]
            }));
            assert!(trailing.is_empty());
            assert_eq!(leading.len(), 1);
            assert!(leading[0].leading);
        }
    }

    #[test]
    fn a_generated_comment_reference_leads_a_break_only_page_and_column_paragraph() {
        let doc = seed_body(&[
            json!({"type": "paragraph", "content": [run("prev")]}),
            json!({"type": "paragraph", "content": [{"type": "run", "content": [
                {"type": "break", "breakType": "page"}, {"type": "break", "breakType": "column"}
            ]}]}),
            json!({"type": "paragraph", "content": [run("next")]}),
        ]);
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "2".into(),
                ranges: vec![(0, 5)],
                author: "A".into(),
                date: String::new(),
                body: Any::Null,
            }],
            &crate::EditCtx::system("2026-09-30"),
        )
        .unwrap();
        let blocks = crate::bridge::yrs_doc_to_layout_blocks(
            &doc,
            "body",
            &crate::bridge::RenderEnv::default(),
        )
        .unwrap();
        let page = blocks
            .iter()
            .find_map(|block| match block {
                docx_layout::types::LayoutBlock::PageBreak(page) => Some(page),
                _ => None,
            })
            .unwrap();
        assert_eq!(page.keeps_leading_spacing, Some(true));
    }

    #[test]
    fn every_open_comment_covers_its_units() {
        let start = |id: u32| json!({"type": "commentRangeStart", "id": id});
        let end = |id: u32| json!({"type": "commentRangeEnd", "id": id});
        let mark =
            |id: u32| json!({"type": "run", "content": [{"type": "commentReference", "id": id}]});
        let cell = json!({"type": "tableCell", "content": [
            {"type": "paragraph", "content": [start(3), run("cell"), end(3), mark(3)]}
        ]});
        // Word's reply layout: the reply's markers beside its parent's.
        let document = seed_body(&[
            json!({"type": "paragraph", "content": [
                start(0), start(1), run("alpha"), end(0), mark(0), end(1), mark(1), start(2), run(" beta")
            ]}),
            json!({"type": "paragraph", "content": [run("gamma"), end(2), mark(2)]}),
            json!({"type": "table", "rows": [{"type": "tableRow", "cells": [cell]}]}),
        ]);
        let covered = |id: &str| {
            document
                .resolve_comment(id)
                .unwrap()
                .into_iter()
                .map(|anchor| (anchor.story, anchor.start, anchor.end))
                .collect::<Vec<_>>()
        };
        let body = |start, end| vec![("body".to_owned(), start, end)];
        assert_eq!(covered("0"), body(0, 5));
        assert_eq!(covered("1"), body(0, 5));
        assert_eq!(covered("2"), body(7, 18));
        assert_eq!(covered("3"), vec![("body:t0:r0c0".to_owned(), 0, 4)]);
    }

    #[test]
    fn reused_run_units_keep_comments_out_of_saved_boundaries() {
        let (units, ppr) = paragraph_units(
            &json!({"content": [
                {"type": "commentRangeStart", "id": 7},
                {"type": "run", "formatting": {"bold": true}, "propertyChanges": [{"id": 3}], "content": [
                    {"type": "text", "text": "A"},
                    {"type": "tab"},
                    {"type": "softHyphen"}
                ]},
                {"type": "commentRangeEnd", "id": 7},
                {"type": "run", "content": [{"type": "footnoteRef", "id": 12}]}
            ]}),
            &StyleResolver::new(None),
            None,
            &BTreeMap::new(),
            &mut OpenComments::default(),
            0,
        );
        assert_eq!(units.len(), 4);
        assert!(units[..3].iter().all(|unit| unit.comments == ["7"]));
        assert!(units[3].comments.is_empty());
        assert_eq!(units[3].pm_size, 1);
        let boundaries = ppr["_originalRunBoundaries"].as_array().unwrap();
        assert_eq!(boundaries.len(), 2);
        assert_eq!(boundaries[0]["text"], "A\t\u{00ad}");
        assert_eq!(
            boundaries[0]["marksKey"],
            "bold:{}|rPrChange:{\"changes\":[{\"currentAttributes\":{},\"id\":3,\"previousAttributes\":{}}]}"
        );
        assert_eq!(boundaries[1]["text"], "12");
    }

    #[test]
    fn multi_digit_note_references_seed_one_position() {
        let styles = StyleResolver::new(None);
        let units = run_to_units(
            &json!({
                "type": "run",
                "content": [{ "type": "footnoteRef", "id": 12 }],
            }),
            None,
            &styles,
            &[],
            &BTreeMap::new(),
        );
        assert_eq!(units.len(), 1);
        assert_eq!(units[0].pm_size, 1);
    }

    #[test]
    fn widow_control_left_unauthored_anywhere_seeds_null() {
        let styles = StyleResolver::new(Some(&json!({
            "styles": [{ "styleId": "Normal", "type": "paragraph", "default": true, "pPr": {} }]
        })));

        assert_eq!(seeded_widow_control(&styles, json!({})), Some(Value::Null));
    }
}
