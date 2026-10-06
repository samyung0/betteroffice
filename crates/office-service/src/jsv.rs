//! JavaScript values with reference semantics, for code ported line by line
//! from TypeScript that relies on them: objects and arrays are shared by
//! reference (`clone` aliases, [`Obj::spread`] copies), `undefined` is a value
//! an object can hold, and object identity is observable ([`V::same`]).
//! Objects enumerate array-index keys first, then other keys in insertion order.

use std::cell::{Ref, RefCell};
use std::rc::Rc;

use indexmap::IndexMap;
use serde::de::{self, Deserialize, Deserializer, MapAccess, SeqAccess, Visitor};
use serde::ser::{Serialize, SerializeMap, SerializeSeq, Serializer};

use crate::js::{number_to_string, quote};

pub type Str = Rc<str>;

#[derive(Clone, Default)]
pub enum V {
    #[default]
    Undef,
    Null,
    Bool(bool),
    Num(f64),
    Str(Str),
    Arr(Arr),
    Obj(Obj),
}

#[derive(Clone, Default)]
pub struct Arr(Rc<RefCell<Vec<V>>>);

#[derive(Clone, Default)]
pub struct Obj(Rc<RefCell<IndexMap<Str, V>>>);

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
        Self::default()
    }

    pub fn same(&self, other: &Obj) -> bool {
        Rc::ptr_eq(&self.0, &other.0)
    }

    /// Identity for maps keyed by object.
    pub fn id(&self) -> usize {
        Rc::as_ptr(&self.0) as *const () as usize
    }

    /// `object[key]`, `undefined` when absent.
    pub fn get(&self, key: &str) -> V {
        self.0.borrow().get(key).cloned().unwrap_or_default()
    }

    /// `key in object`.
    pub fn has(&self, key: &str) -> bool {
        self.0.borrow().contains_key(key)
    }

    /// `object[key] = value`; an existing key keeps its place.
    pub fn set(&self, key: &str, value: V) {
        let mut map = self.0.borrow_mut();
        if let Some(slot) = map.get_mut(key) {
            *slot = value;
            return;
        }
        match array_index(key) {
            Some(index) => {
                let at = map
                    .keys()
                    .position(|other| array_index(other).is_none_or(|other| other > index))
                    .unwrap_or(map.len());
                map.shift_insert(at, Str::from(key), value);
            }
            None => {
                map.insert(Str::from(key), value);
            }
        }
    }

    /// `delete object[key]`.
    pub fn delete(&self, key: &str) {
        self.0.borrow_mut().shift_remove(key);
    }

    /// `Object.keys(object)`.
    pub fn keys(&self) -> Vec<Str> {
        self.0.borrow().keys().cloned().collect()
    }

    /// `Object.entries(object)`.
    pub fn entries(&self) -> Vec<(Str, V)> {
        self.0
            .borrow()
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect()
    }

    /// `Object.values(object)`.
    pub fn values(&self) -> Vec<V> {
        self.0.borrow().values().cloned().collect()
    }

    pub fn len(&self) -> usize {
        self.0.borrow().len()
    }

    pub fn is_empty(&self) -> bool {
        self.0.borrow().is_empty()
    }

    /// `{ ...object }`.
    pub fn spread(&self) -> Obj {
        let copy = Obj::new();
        copy.assign(self);
        copy
    }

    /// `Object.assign(this, other)` / `{ ...this, ...other }` onto this.
    pub fn assign(&self, other: &Obj) {
        for (key, value) in other.entries() {
            self.set(&key, value);
        }
    }

    /// A new object `{ ...this, key: value, ... }`.
    pub fn with(&self, entries: &[(&str, V)]) -> Obj {
        let copy = self.spread();
        for (key, value) in entries {
            copy.set(key, value.clone());
        }
        copy
    }

    pub fn borrow(&self) -> Ref<'_, IndexMap<Str, V>> {
        self.0.borrow()
    }
}

impl Arr {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn from(items: Vec<V>) -> Self {
        Self(Rc::new(RefCell::new(items)))
    }

    pub fn same(&self, other: &Arr) -> bool {
        Rc::ptr_eq(&self.0, &other.0)
    }

    pub fn id(&self) -> usize {
        Rc::as_ptr(&self.0) as *const () as usize
    }

    pub fn len(&self) -> usize {
        self.0.borrow().len()
    }

    pub fn is_empty(&self) -> bool {
        self.0.borrow().is_empty()
    }

    /// `array[index]`.
    pub fn get(&self, index: usize) -> V {
        self.0.borrow().get(index).cloned().unwrap_or_default()
    }

    pub fn set(&self, index: usize, value: V) {
        let mut items = self.0.borrow_mut();
        if index >= items.len() {
            items.resize(index + 1, V::Undef);
        }
        items[index] = value;
    }

    pub fn push(&self, value: V) {
        self.0.borrow_mut().push(value);
    }

    pub fn insert(&self, index: usize, value: V) {
        let mut items = self.0.borrow_mut();
        let index = index.min(items.len());
        items.insert(index, value);
    }

    /// The elements (shared, as a JS loop sees them).
    pub fn items(&self) -> Vec<V> {
        self.0.borrow().clone()
    }

    /// `array.at(-1)`.
    pub fn last(&self) -> V {
        self.0.borrow().last().cloned().unwrap_or_default()
    }

    /// `[...array]`.
    pub fn copy(&self) -> Arr {
        Arr::from(self.items())
    }

    /// `array.slice(start, end)` with non-negative bounds.
    pub fn slice(&self, start: usize, end: usize) -> Arr {
        let items = self.0.borrow();
        let end = end.min(items.len());
        let start = start.min(end);
        Arr::from(items[start..end].to_vec())
    }

    pub fn replace(&self, items: Vec<V>) {
        *self.0.borrow_mut() = items;
    }

    pub fn borrow(&self) -> Ref<'_, Vec<V>> {
        self.0.borrow()
    }
}

impl From<Obj> for V {
    fn from(object: Obj) -> Self {
        V::Obj(object)
    }
}

impl From<Arr> for V {
    fn from(array: Arr) -> Self {
        V::Arr(array)
    }
}

impl From<&str> for V {
    fn from(text: &str) -> Self {
        V::Str(Str::from(text))
    }
}

impl From<String> for V {
    fn from(text: String) -> Self {
        V::Str(Str::from(text))
    }
}

impl From<Str> for V {
    fn from(text: Str) -> Self {
        V::Str(text)
    }
}

impl From<f64> for V {
    fn from(value: f64) -> Self {
        V::Num(value)
    }
}

impl From<bool> for V {
    fn from(value: bool) -> Self {
        V::Bool(value)
    }
}

impl<T: Into<V>> From<Option<T>> for V {
    fn from(value: Option<T>) -> Self {
        value.map(Into::into).unwrap_or(V::Undef)
    }
}

/// `{ key: value, ... }`.
#[macro_export]
macro_rules! obj {
    ($($key:literal : $value:expr),* $(,)?) => {{
        let object = $crate::jsv::Obj::new();
        $(object.set($key, $crate::jsv::V::from($value));)*
        $crate::jsv::V::Obj(object)
    }};
}

/// `[a, b, ...]`.
#[macro_export]
macro_rules! arr {
    ($($value:expr),* $(,)?) => {
        $crate::jsv::V::Arr($crate::jsv::Arr::from(vec![$($crate::jsv::V::from($value)),*]))
    };
}

impl V {
    pub fn str(text: impl Into<Str>) -> V {
        V::Str(text.into())
    }

    pub fn new_obj() -> (V, Obj) {
        let object = Obj::new();
        (V::Obj(object.clone()), object)
    }

    pub fn array(items: Vec<V>) -> V {
        V::Arr(Arr::from(items))
    }

    /// `value?.[key]`: undefined unless an object (arrays read no keys here).
    pub fn get(&self, key: &str) -> V {
        match self {
            V::Obj(object) => object.get(key),
            _ => V::Undef,
        }
    }

    /// `asObject`: a non-null, non-array object.
    pub fn obj(&self) -> Option<&Obj> {
        match self {
            V::Obj(object) => Some(object),
            _ => None,
        }
    }

    /// `Array.isArray(value)`.
    pub fn arr(&self) -> Option<&Arr> {
        match self {
            V::Arr(array) => Some(array),
            _ => None,
        }
    }

    /// The elements of an array, none otherwise.
    pub fn items(&self) -> Vec<V> {
        self.arr().map(Arr::items).unwrap_or_default()
    }

    /// `typeof value === 'string'`.
    pub fn as_str(&self) -> Option<Str> {
        match self {
            V::Str(text) => Some(text.clone()),
            _ => None,
        }
    }

    /// `typeof value === 'number'`.
    pub fn as_num(&self) -> Option<f64> {
        match self {
            V::Num(value) => Some(*value),
            _ => None,
        }
    }

    /// `asFiniteNumber`.
    pub fn finite(&self) -> Option<f64> {
        self.as_num().filter(|value| value.is_finite())
    }

    pub fn as_bool(&self) -> Option<bool> {
        match self {
            V::Bool(value) => Some(*value),
            _ => None,
        }
    }

    pub fn is_undef(&self) -> bool {
        matches!(self, V::Undef)
    }

    /// `value == null`.
    pub fn nullish(&self) -> bool {
        matches!(self, V::Undef | V::Null)
    }

    pub fn truthy(&self) -> bool {
        match self {
            V::Undef | V::Null => false,
            V::Bool(value) => *value,
            V::Num(value) => *value != 0.0 && !value.is_nan(),
            V::Str(text) => !text.is_empty(),
            V::Arr(_) | V::Obj(_) => true,
        }
    }

    /// `a ?? b`.
    pub fn or_else(self, other: impl FnOnce() -> V) -> V {
        if self.nullish() { other() } else { self }
    }

    /// `a || b`.
    pub fn or(self, other: impl FnOnce() -> V) -> V {
        if self.truthy() { self } else { other() }
    }

    /// `a === b`.
    pub fn same(&self, other: &V) -> bool {
        match (self, other) {
            (V::Undef, V::Undef) | (V::Null, V::Null) => true,
            (V::Bool(a), V::Bool(b)) => a == b,
            (V::Num(a), V::Num(b)) => a == b,
            (V::Str(a), V::Str(b)) => a == b,
            (V::Arr(a), V::Arr(b)) => a.same(b),
            (V::Obj(a), V::Obj(b)) => a.same(b),
            _ => false,
        }
    }

    pub fn type_of(&self) -> &'static str {
        match self {
            V::Undef => "undefined",
            V::Null | V::Arr(_) | V::Obj(_) => "object",
            V::Bool(_) => "boolean",
            V::Num(_) => "number",
            V::Str(_) => "string",
        }
    }

    /// `String(value)` / `${value}`.
    pub fn to_js_string(&self) -> String {
        match self {
            V::Undef => "undefined".into(),
            V::Null => "null".into(),
            V::Bool(value) => value.to_string(),
            V::Num(value) => number_to_string(*value),
            V::Str(text) => text.to_string(),
            V::Arr(items) => items
                .items()
                .iter()
                .map(|item| {
                    if item.nullish() {
                        String::new()
                    } else {
                        item.to_js_string()
                    }
                })
                .collect::<Vec<_>>()
                .join(","),
            V::Obj(_) => "[object Object]".into(),
        }
    }

    /// `Number(value)` for the values the ported code converts.
    pub fn to_number(&self) -> f64 {
        match self {
            V::Undef => f64::NAN,
            V::Null => 0.0,
            V::Bool(value) => f64::from(u8::from(*value)),
            V::Num(value) => *value,
            V::Str(text) => {
                let text = text.trim();
                if text.is_empty() {
                    0.0
                } else if let Some(hex) =
                    text.strip_prefix("0x").or_else(|| text.strip_prefix("0X"))
                {
                    u64::from_str_radix(hex, 16)
                        .map(|value| value as f64)
                        .unwrap_or(f64::NAN)
                } else if text == "Infinity" || text == "+Infinity" {
                    f64::INFINITY
                } else if text == "-Infinity" {
                    f64::NEG_INFINITY
                } else if text.bytes().all(|byte| {
                    byte.is_ascii_digit() || matches!(byte, b'.' | b'e' | b'E' | b'+' | b'-')
                }) {
                    text.parse().unwrap_or(f64::NAN)
                } else {
                    f64::NAN
                }
            }
            V::Arr(items) => {
                let items = items.items();
                match items.len() {
                    0 => 0.0,
                    1 => V::str(items[0].to_js_string()).to_number(),
                    _ => f64::NAN,
                }
            }
            V::Obj(_) => f64::NAN,
        }
    }

    pub fn parse(text: &str) -> Result<V, String> {
        serde_json::from_str(text).map_err(|error| error.to_string())
    }

    /// A deep copy (JSON.parse(JSON.stringify(value)) without the text).
    pub fn deep_copy(&self) -> V {
        match self {
            V::Arr(items) => V::array(items.items().iter().map(V::deep_copy).collect()),
            V::Obj(object) => {
                let copy = Obj::new();
                for (key, value) in object.entries() {
                    copy.set(&key, value.deep_copy());
                }
                V::Obj(copy)
            }
            other => other.clone(),
        }
    }
}

impl From<&serde_json::Value> for V {
    fn from(value: &serde_json::Value) -> Self {
        match value {
            serde_json::Value::Null => V::Null,
            serde_json::Value::Bool(value) => V::Bool(*value),
            serde_json::Value::Number(number) => V::Num(number.as_f64().unwrap_or(f64::NAN)),
            serde_json::Value::String(text) => V::str(text.as_str()),
            serde_json::Value::Array(items) => V::array(items.iter().map(V::from).collect()),
            serde_json::Value::Object(map) => {
                let object = Obj::new();
                for (key, value) in map {
                    object.set(key, V::from(value));
                }
                V::Obj(object)
            }
        }
    }
}

impl<'de> Deserialize<'de> for V {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        struct VVisitor;
        impl<'de> Visitor<'de> for VVisitor {
            type Value = V;
            fn expecting(&self, formatter: &mut std::fmt::Formatter) -> std::fmt::Result {
                formatter.write_str("a JSON value")
            }
            fn visit_bool<E: de::Error>(self, value: bool) -> Result<V, E> {
                Ok(V::Bool(value))
            }
            fn visit_i64<E: de::Error>(self, value: i64) -> Result<V, E> {
                Ok(V::Num(value as f64))
            }
            fn visit_u64<E: de::Error>(self, value: u64) -> Result<V, E> {
                Ok(V::Num(value as f64))
            }
            fn visit_f64<E: de::Error>(self, value: f64) -> Result<V, E> {
                Ok(V::Num(value))
            }
            fn visit_str<E: de::Error>(self, value: &str) -> Result<V, E> {
                Ok(V::str(value))
            }
            fn visit_unit<E: de::Error>(self) -> Result<V, E> {
                Ok(V::Null)
            }
            fn visit_none<E: de::Error>(self) -> Result<V, E> {
                Ok(V::Null)
            }
            fn visit_some<D: Deserializer<'de>>(self, deserializer: D) -> Result<V, D::Error> {
                V::deserialize(deserializer)
            }
            fn visit_seq<A: SeqAccess<'de>>(self, mut seq: A) -> Result<V, A::Error> {
                let mut items = Vec::with_capacity(seq.size_hint().unwrap_or(0));
                while let Some(item) = seq.next_element()? {
                    items.push(item);
                }
                Ok(V::array(items))
            }
            fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<V, A::Error> {
                let object = Obj::new();
                while let Some((key, value)) = map.next_entry::<String, V>()? {
                    object.set(&key, value);
                }
                Ok(V::Obj(object))
            }
        }
        deserializer.deserialize_any(VVisitor)
    }
}

/// `JSON.stringify` semantics for serde: `undefined` properties are left
/// out, `undefined` elements and non-finite numbers are null, and integral
/// numbers serialize as integers so typed structs read them.
impl Serialize for V {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        match self {
            V::Undef | V::Null => serializer.serialize_unit(),
            V::Bool(value) => serializer.serialize_bool(*value),
            V::Num(value) => {
                if !value.is_finite() {
                    serializer.serialize_unit()
                } else if value.fract() == 0.0 && value.abs() < 9_007_199_254_740_992.0 {
                    serializer.serialize_i64(*value as i64)
                } else {
                    serializer.serialize_f64(*value)
                }
            }
            V::Str(text) => serializer.serialize_str(text),
            V::Arr(items) => {
                let items = items.borrow();
                let mut seq = serializer.serialize_seq(Some(items.len()))?;
                for item in items.iter() {
                    seq.serialize_element(item)?;
                }
                seq.end()
            }
            V::Obj(object) => {
                let map = object.borrow();
                let mut out = serializer.serialize_map(None)?;
                for (key, value) in map.iter() {
                    if !value.is_undef() {
                        out.serialize_entry(&**key, value)?;
                    }
                }
                out.end()
            }
        }
    }
}

/// `JSON.stringify(value)`; none for a top-level `undefined`.
pub fn stringify(value: &V) -> Option<String> {
    if value.is_undef() {
        return None;
    }
    let mut out = String::new();
    write(&mut out, value);
    Some(out)
}

fn write(out: &mut String, value: &V) {
    match value {
        V::Undef | V::Null => out.push_str("null"),
        V::Bool(true) => out.push_str("true"),
        V::Bool(false) => out.push_str("false"),
        V::Num(number) => {
            if number.is_finite() {
                out.push_str(&number_to_string(*number));
            } else {
                out.push_str("null");
            }
        }
        V::Str(text) => quote(out, text),
        V::Arr(items) => {
            out.push('[');
            for (index, item) in items.borrow().iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                write(out, item);
            }
            out.push(']');
        }
        V::Obj(object) => {
            out.push('{');
            let mut first = true;
            for (key, item) in object.borrow().iter() {
                if item.is_undef() {
                    continue;
                }
                if !first {
                    out.push(',');
                }
                first = false;
                quote(out, key);
                out.push(':');
                write(out, item);
            }
            out.push('}');
        }
    }
}

/// UTF-16 helpers for string positions the TS counts in code units.
pub mod text {
    /// `text.length`.
    pub fn len(text: &str) -> usize {
        text.encode_utf16().count()
    }

    /// `text.slice(start, end)` in code units (a split surrogate pair is lost).
    pub fn slice(text: &str, start: usize, end: usize) -> String {
        let units: Vec<u16> = text.encode_utf16().collect();
        let end = end.min(units.len());
        let start = start.min(end);
        String::from_utf16_lossy(&units[start..end])
    }

    /// `text.slice(start)`.
    pub fn slice_from(text: &str, start: usize) -> String {
        slice(text, start, usize::MAX)
    }
}

impl std::fmt::Debug for V {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match stringify(self) {
            Some(json) => formatter.write_str(&json),
            None => formatter.write_str("undefined"),
        }
    }
}

/// `JSON.stringify(key)` for building keys.
pub fn quoted(text: &str) -> String {
    let mut out = String::new();
    quote(&mut out, text);
    out
}

/// office-checkpoint.ts `canonical`: `JSON.stringify` with every object's
/// keys sorted by `localeCompare` (array-index keys still lead).
pub fn canonical(value: &V) -> String {
    let mut out = String::new();
    write_canonical(&mut out, value);
    out
}

fn write_canonical(out: &mut String, value: &V) {
    match value {
        V::Arr(items) => {
            out.push('[');
            for (index, item) in items.borrow().iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                write_canonical(out, item);
            }
            out.push(']');
        }
        V::Obj(object) => {
            let mut entries: Vec<(Str, V)> = object
                .entries()
                .into_iter()
                .filter(|(_, item)| !item.is_undef())
                .collect();
            entries.sort_by(|(a, _), (b, _)| match (array_index(a), array_index(b)) {
                (Some(a), Some(b)) => a.cmp(&b),
                (Some(_), None) => std::cmp::Ordering::Less,
                (None, Some(_)) => std::cmp::Ordering::Greater,
                (None, None) => crate::js::locale_compare(a, b),
            });
            out.push('{');
            for (index, (key, item)) in entries.iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                quote(out, key);
                out.push(':');
                write_canonical(out, item);
            }
            out.push('}');
        }
        other => write(out, other),
    }
}
