//! JSON values with JavaScript semantics: objects enumerate integer keys first
//! in ascending order and other keys in insertion order (as `JSON.parse` and
//! object literals do), numbers are doubles, and [`stringify`] writes what
//! `JSON.stringify` writes.

use std::cmp::Ordering;
use std::fmt::Write as _;

use indexmap::IndexMap;
use serde::de::{self, Deserialize, Deserializer, MapAccess, SeqAccess, Visitor};
use serde::ser::{Serialize, SerializeMap, SerializeSeq, Serializer};

#[derive(Clone, Debug, Default, PartialEq)]
pub enum J {
    #[default]
    Null,
    Bool(bool),
    Num(f64),
    Str(String),
    Arr(Vec<J>),
    Obj(Obj),
}

/// A JavaScript object's own enumerable string properties.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Obj(IndexMap<String, J>);

/// An array index key (`"0"` to `"4294967294"`), which objects enumerate first.
fn array_index(key: &str) -> Option<u32> {
    if key.is_empty() || key.len() > 10 || (key.len() > 1 && key.starts_with('0')) {
        return None;
    }
    if !key.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    key.parse::<u64>()
        .ok()
        .filter(|value| *value < u32::MAX as u64)
        .map(|value| value as u32)
}

impl Obj {
    pub fn new() -> Self {
        Self(IndexMap::new())
    }

    pub fn get(&self, key: &str) -> Option<&J> {
        self.0.get(key)
    }

    pub fn get_mut(&mut self, key: &str) -> Option<&mut J> {
        self.0.get_mut(key)
    }

    pub fn contains(&self, key: &str) -> bool {
        self.0.contains_key(key)
    }

    /// `object[key] = value`: an existing key keeps its place.
    pub fn set(&mut self, key: impl Into<String>, value: J) {
        let key = key.into();
        if let Some(slot) = self.0.get_mut(&key) {
            *slot = value;
            return;
        }
        match array_index(&key) {
            Some(index) => {
                let at = self
                    .0
                    .keys()
                    .position(|other| array_index(other).is_none_or(|other| other > index))
                    .unwrap_or(self.0.len());
                self.0.shift_insert(at, key, value);
            }
            None => {
                self.0.insert(key, value);
            }
        }
    }

    /// `delete object[key]`.
    pub fn remove(&mut self, key: &str) -> Option<J> {
        self.0.shift_remove(key)
    }

    pub fn iter(&self) -> impl Iterator<Item = (&String, &J)> {
        self.0.iter()
    }

    pub fn keys(&self) -> impl Iterator<Item = &String> {
        self.0.keys()
    }

    pub fn len(&self) -> usize {
        self.0.len()
    }

    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

impl FromIterator<(String, J)> for Obj {
    fn from_iter<T: IntoIterator<Item = (String, J)>>(iter: T) -> Self {
        let mut object = Obj::new();
        for (key, value) in iter {
            object.set(key, value);
        }
        object
    }
}

impl J {
    pub fn str(value: impl Into<String>) -> J {
        J::Str(value.into())
    }

    pub fn obj<const N: usize>(entries: [(&str, J); N]) -> J {
        J::Obj(
            entries
                .into_iter()
                .map(|(key, value)| (key.to_owned(), value))
                .collect(),
        )
    }

    pub fn get(&self, key: &str) -> Option<&J> {
        match self {
            J::Obj(object) => object.get(key),
            _ => None,
        }
    }

    pub fn as_obj(&self) -> Option<&Obj> {
        match self {
            J::Obj(object) => Some(object),
            _ => None,
        }
    }

    pub fn as_arr(&self) -> Option<&Vec<J>> {
        match self {
            J::Arr(items) => Some(items),
            _ => None,
        }
    }

    pub fn as_str(&self) -> Option<&str> {
        match self {
            J::Str(text) => Some(text),
            _ => None,
        }
    }

    pub fn as_f64(&self) -> Option<f64> {
        match self {
            J::Num(value) => Some(*value),
            _ => None,
        }
    }

    pub fn is_null(&self) -> bool {
        matches!(self, J::Null)
    }

    /// JavaScript truthiness of a JSON value.
    pub fn truthy(&self) -> bool {
        match self {
            J::Null => false,
            J::Bool(value) => *value,
            J::Num(value) => *value != 0.0 && !value.is_nan(),
            J::Str(text) => !text.is_empty(),
            J::Arr(_) | J::Obj(_) => true,
        }
    }

    pub fn parse(text: &str) -> Result<J, String> {
        serde_json::from_str(text).map_err(|error| error.to_string())
    }
}

impl<'de> Deserialize<'de> for J {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        struct JVisitor;
        impl<'de> Visitor<'de> for JVisitor {
            type Value = J;
            fn expecting(&self, formatter: &mut std::fmt::Formatter) -> std::fmt::Result {
                formatter.write_str("a JSON value")
            }
            fn visit_bool<E: de::Error>(self, value: bool) -> Result<J, E> {
                Ok(J::Bool(value))
            }
            fn visit_i64<E: de::Error>(self, value: i64) -> Result<J, E> {
                Ok(J::Num(value as f64))
            }
            fn visit_u64<E: de::Error>(self, value: u64) -> Result<J, E> {
                Ok(J::Num(value as f64))
            }
            fn visit_f64<E: de::Error>(self, value: f64) -> Result<J, E> {
                Ok(J::Num(value))
            }
            fn visit_str<E: de::Error>(self, value: &str) -> Result<J, E> {
                Ok(J::Str(value.to_owned()))
            }
            fn visit_string<E: de::Error>(self, value: String) -> Result<J, E> {
                Ok(J::Str(value))
            }
            fn visit_unit<E: de::Error>(self) -> Result<J, E> {
                Ok(J::Null)
            }
            fn visit_none<E: de::Error>(self) -> Result<J, E> {
                Ok(J::Null)
            }
            fn visit_some<D: Deserializer<'de>>(self, deserializer: D) -> Result<J, D::Error> {
                J::deserialize(deserializer)
            }
            fn visit_seq<A: SeqAccess<'de>>(self, mut seq: A) -> Result<J, A::Error> {
                let mut items = Vec::with_capacity(seq.size_hint().unwrap_or(0));
                while let Some(item) = seq.next_element()? {
                    items.push(item);
                }
                Ok(J::Arr(items))
            }
            fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<J, A::Error> {
                let mut object = Obj::new();
                while let Some((key, value)) = map.next_entry::<String, J>()? {
                    object.set(key, value);
                }
                Ok(J::Obj(object))
            }
        }
        deserializer.deserialize_any(JVisitor)
    }
}

/// Integral doubles serialize as integers, so typed engine structs read them.
impl Serialize for J {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        match self {
            J::Null => serializer.serialize_unit(),
            J::Bool(value) => serializer.serialize_bool(*value),
            J::Num(value) => {
                if !value.is_finite() {
                    serializer.serialize_unit()
                } else if value.fract() == 0.0 && value.abs() < 9_007_199_254_740_992.0 {
                    serializer.serialize_i64(*value as i64)
                } else {
                    serializer.serialize_f64(*value)
                }
            }
            J::Str(text) => serializer.serialize_str(text),
            J::Arr(items) => {
                let mut seq = serializer.serialize_seq(Some(items.len()))?;
                for item in items {
                    seq.serialize_element(item)?;
                }
                seq.end()
            }
            J::Obj(object) => {
                let mut map = serializer.serialize_map(Some(object.len()))?;
                for (key, value) in object.iter() {
                    map.serialize_entry(key, value)?;
                }
                map.end()
            }
        }
    }
}

impl From<serde_json::Value> for J {
    fn from(value: serde_json::Value) -> Self {
        match value {
            serde_json::Value::Null => J::Null,
            serde_json::Value::Bool(value) => J::Bool(value),
            serde_json::Value::Number(number) => J::Num(number.as_f64().unwrap_or(f64::NAN)),
            serde_json::Value::String(text) => J::Str(text),
            serde_json::Value::Array(items) => J::Arr(items.into_iter().map(J::from).collect()),
            serde_json::Value::Object(map) => J::Obj(
                map.into_iter()
                    .map(|(key, value)| (key, J::from(value)))
                    .collect(),
            ),
        }
    }
}

/// `String(number)` / the number in `JSON.stringify`.
pub fn number_to_string(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    let mut buffer = ryu_js::Buffer::new();
    buffer.format(value).to_owned()
}

/// `JSON.stringify(text)`.
pub fn quote(out: &mut String, text: &str) {
    out.push('"');
    for ch in text.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\u{8}' => out.push_str("\\b"),
            '\u{c}' => out.push_str("\\f"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            ch if (ch as u32) < 0x20 => {
                let _ = write!(out, "\\u{:04x}", ch as u32);
            }
            ch => out.push(ch),
        }
    }
    out.push('"');
}

/// `JSON.stringify(value)` without a replacer.
pub fn stringify(value: &J) -> String {
    let mut out = String::new();
    write_json(&mut out, value, false);
    out
}

/// office-checkpoint.ts `canonical`: `JSON.stringify` with every object's
/// keys sorted by `localeCompare` (array-index keys still enumerate first).
pub fn canonical(value: &J) -> String {
    let mut out = String::new();
    write_json(&mut out, value, true);
    out
}

fn write_json(out: &mut String, value: &J, sorted: bool) {
    match value {
        J::Null => out.push_str("null"),
        J::Bool(true) => out.push_str("true"),
        J::Bool(false) => out.push_str("false"),
        J::Num(number) => {
            if number.is_finite() {
                out.push_str(&number_to_string(*number));
            } else {
                out.push_str("null");
            }
        }
        J::Str(text) => quote(out, text),
        J::Arr(items) => {
            out.push('[');
            for (index, item) in items.iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                write_json(out, item, sorted);
            }
            out.push(']');
        }
        J::Obj(object) => {
            out.push('{');
            let mut entries: Vec<(&String, &J)> = object.iter().collect();
            if sorted {
                // Object.fromEntries of the sorted entries: index keys still lead.
                entries.sort_by(|(a, _), (b, _)| match (array_index(a), array_index(b)) {
                    (Some(a), Some(b)) => a.cmp(&b),
                    (Some(_), None) => Ordering::Less,
                    (None, Some(_)) => Ordering::Greater,
                    (None, None) => locale_compare(a, b),
                });
            }
            for (index, (key, item)) in entries.into_iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                quote(out, key);
                out.push(':');
                write_json(out, item, sorted);
            }
            out.push('}');
        }
    }
}

/// Primary weight of an ASCII character in the CLDR root collation (what
/// `String.prototype.localeCompare` uses with the default locale): variable
/// characters, then digits, then letters case-insensitively.
fn primary(ch: char) -> u32 {
    const VARIABLE: &str = "\t\n\u{b}\u{c}\r _-,;:!?.'\"()[]{}@*/\\&#%`^+<=>|~$";
    if let Some(index) = VARIABLE.find(ch) {
        return 1 + index as u32;
    }
    match ch {
        '0'..='9' => 100 + (ch as u32 - '0' as u32),
        'a'..='z' => 200 + (ch as u32 - 'a' as u32),
        'A'..='Z' => 200 + (ch as u32 - 'A' as u32),
        // Outside ASCII: code point order after every ASCII character.
        _ => 1000 + ch as u32,
    }
}

/// `a.localeCompare(b)` for the keys these documents hold: primary weights
/// first, then lowercase before uppercase at the first case difference.
pub fn locale_compare(a: &str, b: &str) -> Ordering {
    let primaries = |text: &str| text.chars().map(primary).collect::<Vec<_>>();
    match primaries(a).cmp(&primaries(b)) {
        Ordering::Equal => {}
        other => return other,
    }
    for (x, y) in a.chars().zip(b.chars()) {
        if x != y {
            let rank = |ch: char| u8::from(ch.is_ascii_uppercase());
            return rank(x).cmp(&rank(y)).then(x.cmp(&y));
        }
    }
    a.len().cmp(&b.len())
}

/// Default `Array.prototype.sort` order: UTF-16 code units.
pub fn utf16_compare(a: &str, b: &str) -> Ordering {
    a.encode_utf16().cmp(b.encode_utf16())
}

/// `text.length`.
pub fn utf16_len(text: &str) -> usize {
    text.encode_utf16().count()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn objects_enumerate_index_keys_first() {
        let value = J::parse(r#"{"b":1,"10":2,"a":3,"2":4}"#).unwrap();
        assert_eq!(stringify(&value), r#"{"2":4,"10":2,"b":1,"a":3}"#);
        assert_eq!(canonical(&value), r#"{"2":4,"10":2,"a":3,"b":1}"#);
    }

    #[test]
    fn numbers_and_strings_print_as_javascript_does() {
        let value = J::parse(r#"[1.0,0.5,1e21,-0.0,1e-7,"a\u0001\" "]"#).unwrap();
        assert_eq!(
            stringify(&value),
            "[1,0.5,1e+21,0,1e-7,\"a\\u0001\\\"\u{2028}\"]"
        );
    }

    #[test]
    fn keys_sort_as_locale_compare_does() {
        let mut keys = vec![
            "colWidths",
            "_x",
            "Zeta",
            "alpha",
            "colwidths",
            "a1",
            "a_",
            "B",
        ];
        keys.sort_by(|a, b| locale_compare(a, b));
        assert_eq!(
            keys,
            [
                "_x",
                "a_",
                "a1",
                "alpha",
                "B",
                "colwidths",
                "colWidths",
                "Zeta"
            ]
        );
    }
}
