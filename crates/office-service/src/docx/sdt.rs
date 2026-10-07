//! Content controls: types/sdtAttributes.ts `sdtAttrsToProps` and
//! yrs/contentControlValues.ts `applyContentControlValue`.

use crate::jsv::{Obj, V, text};
use crate::{arr, obj};

/// Rebuilds structured content-control properties from the flat yrs vocabulary.
pub(crate) fn sdt_attrs_to_props(attrs: &V) -> V {
    let get = |key: &str| attrs.get(key);
    let props = Obj::new();
    props.set("sdtType", get("sdtType").or_else(|| V::str("richText")));
    if get("id").as_num().is_some() {
        props.set("id", get("id"));
    }
    for key in ["alias", "tag"] {
        if !get(key).nullish() {
            props.set(key, V::from(get(key).to_js_string()));
        }
    }
    if !get("lock").nullish() {
        props.set("lock", get("lock"));
    }
    if !get("placeholder").nullish() {
        props.set("placeholder", V::from(get("placeholder").to_js_string()));
    }
    if get("showingPlaceholder").truthy() {
        props.set("showingPlaceholder", V::Bool(true));
    }
    if !get("dateFormat").nullish() {
        props.set("dateFormat", V::from(get("dateFormat").to_js_string()));
    }
    if let Some(items) = get("listItems").as_str().filter(|items| !items.is_empty())
        && let Ok(parsed) = V::parse(&items)
    {
        props.set("listItems", parsed);
    }
    if !get("checked").nullish() {
        props.set("checked", get("checked"));
    }
    if let Some(binding) = get("dataBinding")
        .as_str()
        .filter(|binding| !binding.is_empty())
        && let Ok(parsed) = V::parse(&binding)
    {
        props.set("dataBinding", parsed);
    }
    for key in ["rawPropertiesXml", "rawEndPropertiesXml"] {
        if !get(key).nullish() {
            props.set(key, V::from(get(key).to_js_string()));
        }
    }
    V::Obj(props)
}

/// The first `<element ...>` tag in `xml` (`new RegExp(`<${element}\\b[^>]*>`)`).
fn element_tag<'a>(xml: &'a str, element: &str) -> Option<(usize, &'a str)> {
    let opening = format!("<{element}");
    let mut from = 0;
    while let Some(found) = xml[from..].find(&opening) {
        let start = from + found;
        let after = &xml[start + opening.len()..];
        let boundary = after
            .chars()
            .next()
            .is_none_or(|ch| !(ch.is_ascii_alphanumeric() || ch == '_'));
        if boundary && let Some(end) = after.find('>') {
            return Some((start, &xml[start..start + opening.len() + end + 1]));
        }
        from = start + 1;
    }
    None
}

/// `\battr="value"` in `tag`: the value's byte range.
fn attribute_span(tag: &str, attr: &str) -> Option<(usize, usize)> {
    let needle = format!("{attr}=\"");
    let mut from = 0;
    while let Some(found) = tag[from..].find(&needle) {
        let start = from + found;
        let boundary = tag[..start]
            .chars()
            .last()
            .is_none_or(|ch| !(ch.is_ascii_alphanumeric() || ch == '_'));
        if boundary {
            let value_start = start + needle.len();
            let value_end = value_start + tag[value_start..].find('"')?;
            return Some((start, value_end + 1));
        }
        from = start + 1;
    }
    None
}

fn read_attr(xml: &str, element: &str, attr: &str) -> Option<String> {
    let (_, tag) = element_tag(xml, element)?;
    let (start, end) = attribute_span(tag, attr)?;
    Some(tag[start + attr.len() + 2..end - 1].to_owned())
}

/// Sets or adds an attribute on the first `<element ...>`; unchanged without one.
fn set_attr(xml: &str, element: &str, attr: &str, value: &str) -> String {
    let Some((start, open)) = element_tag(xml, element) else {
        return xml.to_owned();
    };
    let self_close = open.ends_with("/>");
    let body = &open[1..open.len() - if self_close { 2 } else { 1 }];
    let new_body = match attribute_span(body, attr) {
        Some((from, to)) => format!("{}{attr}=\"{value}\"{}", &body[..from], &body[to..]),
        None => format!("{body} {attr}=\"{value}\""),
    };
    let replaced = format!("<{new_body}{}", if self_close { "/>" } else { ">" });
    format!(
        "{}{}{}",
        &xml[..start],
        replaced,
        &xml[start + open.len()..]
    )
}

fn escape_xml_attr(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// `parseInt(hex, 16)` as a character, else `fallback`.
fn code_point_char(hex: Option<&str>, fallback: &str) -> String {
    let Some(hex) = hex.filter(|hex| !hex.is_empty()) else {
        return fallback.to_owned();
    };
    let digits: String = hex
        .trim_start()
        .chars()
        .take_while(char::is_ascii_hexdigit)
        .collect();
    u32::from_str_radix(&digits, 16)
        .ok()
        .filter(|code| !(0xd800..=0xdfff).contains(code))
        .and_then(char::from_u32)
        .map(String::from)
        .unwrap_or_else(|| fallback.to_owned())
}

/// `parseInt(text, 10)`.
fn parse_int(text: &str) -> Option<i64> {
    let text = text.trim_start();
    let (sign, digits) = match text.strip_prefix('-') {
        Some(rest) => (-1, rest),
        None => (1, text.strip_prefix('+').unwrap_or(text)),
    };
    let digits: String = digits.chars().take_while(char::is_ascii_digit).collect();
    digits.parse::<i64>().ok().map(|value| sign * value)
}

fn valid_iso_date(year: Option<i64>, month: Option<i64>, day: Option<i64>) -> bool {
    let (Some(year), Some(month), Some(day)) = (year, month, day) else {
        return false;
    };
    if year == 0 || month == 0 || day == 0 || (0..100).contains(&year) {
        return false;
    }
    let leap = (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
    let days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return false,
    };
    (1..=days).contains(&day)
}

/// `new Intl.DateTimeFormat(language, { month, timeZone: "UTC" }).format(date)`
/// through ICU4X: the month alone, `long` ("March") or short ("Mar"), in the
/// language's calendar. `None` when the language is not a valid tag.
fn month_name(language: &str, (year, month, day): (i64, i64, i64), long: bool) -> Option<String> {
    use icu_datetime::fieldsets::M;
    let locale: icu_locale_core::Locale = language.parse().ok()?;
    let date = icu_calendar::Date::try_new_iso(
        i32::try_from(year).ok()?,
        u8::try_from(month).ok()?,
        u8::try_from(day).ok()?,
    )
    .ok()?;
    let fields = if long { M::long() } else { M::medium() };
    let formatter = icu_datetime::DateTimeFormatter::try_new(locale.into(), fields).ok()?;
    Some(formatter.format(&date).to_string())
}

/// contentControlValues.ts `formatSdtDate`: month names in `language`
/// (English when it is missing or not a valid tag, as the TS's `Intl` throws).
fn format_sdt_date(iso: &str, pattern: Option<&str>, language: Option<&str>) -> String {
    let head = text::slice(iso, 0, 10);
    let parts: Vec<&str> = head.split('-').collect();
    let part = |index: usize| parts.get(index).and_then(|part| parse_int(part));
    let (year, month, day) = (part(0), part(1), part(2));
    if !valid_iso_date(year, month, day) {
        return iso.to_owned();
    }
    let (year, month, day) = (year.unwrap_or(0), month.unwrap_or(1), day.unwrap_or(1));
    let format = pattern
        .filter(|pattern| !pattern.trim().is_empty())
        .unwrap_or("M/d/yyyy");
    let month_in = |long: bool| {
        language
            .filter(|language| !language.is_empty())
            .and_then(|language| month_name(language, (year, month, day), long))
            .or_else(|| month_name("en", (year, month, day), long))
            .unwrap_or_default()
    };
    let year_text = year.to_string();
    let tokens: [(&str, String); 8] = [
        ("yyyy", year_text.clone()),
        (
            "yy",
            year_text[year_text.len().saturating_sub(2)..].to_owned(),
        ),
        ("MMMM", month_in(true)),
        ("MMM", month_in(false)),
        ("MM", format!("{month:02}")),
        ("M", month.to_string()),
        ("dd", format!("{day:02}")),
        ("d", day.to_string()),
    ];
    let mut out = String::new();
    let mut rest = format;
    'scan: while !rest.is_empty() {
        for (token, value) in &tokens {
            if rest.starts_with(token) {
                out.push_str(value);
                rest = &rest[token.len()..];
                continue 'scan;
            }
        }
        let ch = rest.chars().next().unwrap_or_default();
        out.push(ch);
        rest = &rest[ch.len_utf8()..];
    }
    out
}

/// A one-run paragraph; `font` sets the run's font (for symbol glyphs).
fn paragraph(text: V, font: Option<&str>) -> V {
    if !text.truthy() {
        return obj! { "type": "paragraph", "content": arr![] };
    }
    let run = obj! { "type": "run", "content": arr![obj! { "type": "text", "text": text }] };
    if let Some(font) = font.filter(|font| !font.is_empty())
        && let Some(object) = run.obj()
    {
        object.set(
            "formatting",
            obj! { "fontFamily": obj! { "ascii": font, "hAnsi": font, "eastAsia": font, "cs": font } },
        );
    }
    obj! { "type": "paragraph", "content": arr![run] }
}

/// `raw` without `<w:showingPlcHdr .../>`, then without
/// `<w:showingPlcHdr ...>...</w:showingPlcHdr>` (both regexes global).
fn clear_showing_placeholder(raw: &str) -> String {
    const CLOSE: &str = "</w:showingPlcHdr>";
    let mut out = String::new();
    let mut rest = raw;
    while let Some((start, tag)) = element_tag(rest, "w:showingPlcHdr") {
        out.push_str(&rest[..start]);
        if !tag.ends_with("/>") {
            out.push_str(tag);
        }
        rest = &rest[start + tag.len()..];
    }
    out.push_str(rest);
    let mut cleared = String::new();
    let mut rest = out.as_str();
    while let Some((start, tag)) = element_tag(rest, "w:showingPlcHdr") {
        let after = start + tag.len();
        let Some(close) = rest[after..].find(CLOSE) else {
            break;
        };
        cleared.push_str(&rest[..start]);
        rest = &rest[after + close + CLOSE.len()..];
    }
    cleared.push_str(rest);
    cleared
}

fn without_placeholder(props: &Obj, next_raw: &str) -> Obj {
    let cleaned = clear_showing_placeholder(next_raw);
    props.with(&[
        ("showingPlaceholder", V::Bool(false)),
        (
            "rawPropertiesXml",
            if cleaned.is_empty() {
                V::Undef
            } else {
                V::from(cleaned)
            },
        ),
    ])
}

/// The new properties and display blocks for a typed value; `Err` where the
/// TS throws (the projection then keeps the control as it was).
pub(crate) fn apply_content_control_value(props: &V, value: &V) -> Result<(V, Vec<V>), ()> {
    let object = props.obj().cloned().unwrap_or_default();
    let raw = props
        .get("rawPropertiesXml")
        .or_else(|| V::str(""))
        .to_js_string();
    let sdt_type = props.get("sdtType").as_str();
    let control_state = |entries: &[(&str, V)]| {
        let state = props
            .get("controlState")
            .obj()
            .map(Obj::spread)
            .unwrap_or_default();
        for (key, value) in entries {
            state.set(key, value.clone());
        }
        V::Obj(state)
    };
    match value.get("kind").as_str().as_deref() {
        Some("dropdown") => {
            if !matches!(sdt_type.as_deref(), Some("dropDownList" | "comboBox")) {
                return Err(());
            }
            let wanted = value.get("value");
            let items = props.get("listItems");
            let found = items.arr().and_then(|items| {
                items.items().into_iter().enumerate().find(|(_, item)| {
                    item.get("value").same(&wanted) || item.get("displayText").same(&wanted)
                })
            });
            if found.is_none() && sdt_type.as_deref() == Some("dropDownList") {
                return Err(());
            }
            let stored = found
                .as_ref()
                .map(|(_, item)| item.get("value"))
                .unwrap_or_else(|| wanted.clone())
                .or_else(|| wanted.clone());
            let display = found
                .as_ref()
                .map(|(_, item)| item.get("displayText"))
                .unwrap_or_else(|| wanted.clone())
                .or_else(|| wanted.clone());
            let selected = found
                .map(|(index, _)| V::Num(index as f64))
                .unwrap_or(V::Undef);
            let element = if sdt_type.as_deref() == Some("comboBox") {
                "w:comboBox"
            } else {
                "w:dropDownList"
            };
            let Some(stored_text) = stored.as_str() else {
                return Err(());
            };
            let next_raw = set_attr(&raw, element, "w:lastValue", &escape_xml_attr(&stored_text));
            let properties = without_placeholder(&object, &next_raw);
            properties.set("listLastValue", stored.clone());
            properties.set(
                "controlState",
                control_state(&[
                    ("value", display.clone()),
                    ("selectedValue", stored),
                    ("selectedIndex", selected),
                    ("placeholder", V::Bool(false)),
                ]),
            );
            Ok((V::Obj(properties), vec![paragraph(display, None)]))
        }
        Some("checkbox") => {
            if sdt_type.as_deref() != Some("checkbox") {
                return Err(());
            }
            if read_attr(&raw, "w14:checked", "w14:val").is_none() {
                return Err(());
            }
            let checked = value.get("checked").truthy();
            let state_element = if checked {
                "w14:checkedState"
            } else {
                "w14:uncheckedState"
            };
            let glyph = code_point_char(
                read_attr(&raw, state_element, "w14:val").as_deref(),
                if checked { "☒" } else { "☐" },
            );
            let font = read_attr(&raw, state_element, "w14:font");
            let next_raw = set_attr(
                &raw,
                "w14:checked",
                "w14:val",
                if checked { "1" } else { "0" },
            );
            let properties = without_placeholder(&object, &next_raw);
            properties.set("checked", V::Bool(checked));
            properties.set(
                "controlState",
                control_state(&[
                    ("checked", V::Bool(checked)),
                    ("value", V::from(glyph.clone())),
                    ("placeholder", V::Bool(false)),
                ]),
            );
            Ok((
                V::Obj(properties),
                vec![paragraph(V::from(glyph), font.as_deref())],
            ))
        }
        Some("date") => {
            if sdt_type.as_deref() != Some("date") {
                return Err(());
            }
            let iso = text::slice(&value.get("date").to_js_string(), 0, 10);
            let shaped = iso.len() == 10
                && iso.bytes().enumerate().all(|(index, byte)| match index {
                    4 | 7 => byte == b'-',
                    _ => byte.is_ascii_digit(),
                });
            let (year, month, day) = if shaped {
                (
                    iso[0..4].parse().ok(),
                    iso[5..7].parse().ok(),
                    iso[8..10].parse().ok(),
                )
            } else {
                (None, None, None)
            };
            if !valid_iso_date(year, month, day) {
                return Err(());
            }
            let full_date = format!("{iso}T00:00:00");
            let next_raw = set_attr(&raw, "w:date", "w:fullDate", &full_date);
            let pattern = props
                .get("dateState")
                .get("format")
                .or_else(|| props.get("dateFormat"))
                .or_else(|| {
                    read_attr(&raw, "w:dateFormat", "w:val")
                        .map(V::from)
                        .unwrap_or_default()
                });
            let language = props.get("dateState").get("language");
            let display = format_sdt_date(
                &iso,
                pattern.as_str().as_deref(),
                language.as_str().as_deref(),
            );
            let properties = without_placeholder(&object, &next_raw);
            properties.set("dateFormat", pattern.clone());
            let date_state = props
                .get("dateState")
                .obj()
                .map(Obj::spread)
                .unwrap_or_default();
            date_state.set("fullDate", V::from(full_date));
            date_state.set("format", pattern);
            properties.set("dateState", V::Obj(date_state));
            properties.set(
                "controlState",
                control_state(&[("value", V::from(iso)), ("placeholder", V::Bool(false))]),
            );
            Ok((V::Obj(properties), vec![paragraph(V::from(display), None)]))
        }
        _ => Err(()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attributes_patch_like_the_ts_regexes() {
        let raw = r#"<w:sdtPr><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>"#;
        assert_eq!(
            read_attr(raw, "w14:checked", "w14:val").as_deref(),
            Some("0")
        );
        assert_eq!(
            read_attr(raw, "w14:checkedState", "w14:font").as_deref(),
            Some("MS Gothic")
        );
        assert!(
            set_attr(raw, "w14:checked", "w14:val", "1").contains(r#"<w14:checked w14:val="1"/>"#)
        );
        assert_eq!(
            set_attr("<w:date/>", "w:date", "w:fullDate", "2026-01-02T00:00:00"),
            r#"<w:date w:fullDate="2026-01-02T00:00:00"/>"#
        );
        assert_eq!(
            format_sdt_date("2026-03-04", Some("d MMMM yyyy"), None),
            "4 March 2026"
        );
        assert_eq!(code_point_char(Some("2612"), "x"), "☒");
    }

    /// Month names as Node 22.19 `Intl.DateTimeFormat` (ICU 77.1, CLDR 47)
    /// gives them to the TS `formatSdtDate` for the 15th of each 2026 month:
    /// long and short, in the language's calendar; English for a missing or
    /// invalid tag. Generated by `parity/months.mjs` (run with Node).
    #[test]
    fn month_names_match_intl() {
        for (language, long, short) in MONTHS_BY_INTL {
            for (index, (long, short)) in long.split('|').zip(short.split('|')).enumerate() {
                let month = index as i64 + 1;
                let date = format!("2026-{month:02}-15");
                let language = Some(language).filter(|language| !language.is_empty());
                assert_eq!(
                    format_sdt_date(&date, Some("MMMM"), language),
                    long,
                    "{language:?} {month}"
                );
                assert_eq!(
                    format_sdt_date(&date, Some("MMM"), language),
                    short,
                    "{language:?} {month}"
                );
            }
        }
        assert_eq!(
            format_sdt_date("2026-03-04", Some("yyyy年M月d日 (MMMM)"), Some("ja-JP")),
            "2026年3月4日 (3月)"
        );
    }

    const MONTHS_BY_INTL: [(&str, &str, &str); 23] = [
        (
            "en-US",
            "January|February|March|April|May|June|July|August|September|October|November|December",
            "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec",
        ),
        (
            "ja-JP",
            "1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月",
            "1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月",
        ),
        (
            "zh-TW",
            "1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月",
            "1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月",
        ),
        (
            "zh-CN",
            "一月|二月|三月|四月|五月|六月|七月|八月|九月|十月|十一月|十二月",
            "1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月",
        ),
        (
            "zh-HK",
            "1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月",
            "1月|2月|3月|4月|5月|6月|7月|8月|9月|10月|11月|12月",
        ),
        (
            "ko-KR",
            "1월|2월|3월|4월|5월|6월|7월|8월|9월|10월|11월|12월",
            "1월|2월|3월|4월|5월|6월|7월|8월|9월|10월|11월|12월",
        ),
        (
            "fr-FR",
            "janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre",
            "janv.|févr.|mars|avr.|mai|juin|juil.|août|sept.|oct.|nov.|déc.",
        ),
        (
            "de-DE",
            "Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember",
            "Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez",
        ),
        (
            "es-ES",
            "enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre",
            "ene|feb|mar|abr|may|jun|jul|ago|sept|oct|nov|dic",
        ),
        (
            "pt-BR",
            "janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro",
            "jan.|fev.|mar.|abr.|mai.|jun.|jul.|ago.|set.|out.|nov.|dez.",
        ),
        (
            "it-IT",
            "gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre",
            "gen|feb|mar|apr|mag|giu|lug|ago|set|ott|nov|dic",
        ),
        (
            "ru-RU",
            "январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь",
            "янв.|февр.|март|апр.|май|июнь|июль|авг.|сент.|окт.|нояб.|дек.",
        ),
        (
            "vi-VN",
            "Tháng 1|Tháng 2|Tháng 3|Tháng 4|Tháng 5|Tháng 6|Tháng 7|Tháng 8|Tháng 9|Tháng 10|Tháng 11|Tháng 12",
            "Tháng 1|Tháng 2|Tháng 3|Tháng 4|Tháng 5|Tháng 6|Tháng 7|Tháng 8|Tháng 9|Tháng 10|Tháng 11|Tháng 12",
        ),
        (
            "th-TH",
            "มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม",
            "ม.ค.|ก.พ.|มี.ค.|เม.ย.|พ.ค.|มิ.ย.|ก.ค.|ส.ค.|ก.ย.|ต.ค.|พ.ย.|ธ.ค.",
        ),
        (
            "ar-SA",
            "يناير|فبراير|مارس|أبريل|مايو|يونيو|يوليو|أغسطس|سبتمبر|أكتوبر|نوفمبر|ديسمبر",
            "يناير|فبراير|مارس|أبريل|مايو|يونيو|يوليو|أغسطس|سبتمبر|أكتوبر|نوفمبر|ديسمبر",
        ),
        (
            "he-IL",
            "ינואר|פברואר|מרץ|אפריל|מאי|יוני|יולי|אוגוסט|ספטמבר|אוקטובר|נובמבר|דצמבר",
            "ינו׳|פבר׳|מרץ|אפר׳|מאי|יוני|יולי|אוג׳|ספט׳|אוק׳|נוב׳|דצמ׳",
        ),
        (
            "fa-IR",
            "دی|بهمن|اسفند|فروردین|اردیبهشت|خرداد|تیر|مرداد|شهریور|مهر|آبان|آذر",
            "دی|بهمن|اسفند|فروردین|اردیبهشت|خرداد|تیر|مرداد|شهریور|مهر|آبان|آذر",
        ),
        (
            "hi-IN",
            "जनवरी|फ़रवरी|मार्च|अप्रैल|मई|जून|जुलाई|अगस्त|सितंबर|अक्टूबर|नवंबर|दिसंबर",
            "जन॰|फ़र॰|मार्च|अप्रैल|मई|जून|जुल॰|अग॰|सित॰|अक्टू॰|नव॰|दिस॰",
        ),
        (
            "nl-NL",
            "januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december",
            "jan|feb|mrt|apr|mei|jun|jul|aug|sep|okt|nov|dec",
        ),
        (
            "pl-PL",
            "styczeń|luty|marzec|kwiecień|maj|czerwiec|lipiec|sierpień|wrzesień|październik|listopad|grudzień",
            "sty|lut|mar|kwi|maj|cze|lip|sie|wrz|paź|lis|gru",
        ),
        (
            "en_US",
            "January|February|March|April|May|June|July|August|September|October|November|December",
            "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec",
        ),
        (
            "",
            "January|February|March|April|May|June|July|August|September|October|November|December",
            "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec",
        ),
        (
            "not a tag",
            "January|February|March|April|May|June|July|August|September|October|November|December",
            "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec",
        ),
    ];
}
