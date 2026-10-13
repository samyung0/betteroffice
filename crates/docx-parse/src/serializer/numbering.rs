//! S10 `numbering.xml` serializer.

use crate::formatting::ParagraphFormatting;
use crate::numbering::{AbstractNumbering, ListLevel, NumberingDefinitions, NumberingInstance};

use super::xml_writer::{XmlWriter, int_attr};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// Serialize numbering definitions to a complete `word/numbering.xml` part.
pub fn serialize_numbering_xml(numbering: &NumberingDefinitions) -> String {
    let mut writer = XmlWriter::with_capacity(256 + numbering.abstract_nums.len() * 128);
    writer
        .declaration()
        .text("\n")
        .start_element("w:numbering")
        .attribute("xmlns:w", W_NS);
    for abstract_numbering in &numbering.abstract_nums {
        write_abstract_numbering(&mut writer, abstract_numbering);
    }
    for instance in &numbering.nums {
        write_numbering_instance(&mut writer, instance);
    }
    if numbering.abstract_nums.is_empty() && numbering.nums.is_empty() {
        // Empty roots use explicit start and end tags.
        writer.text("");
    }
    writer.end_element();
    writer.finish()
}

/// One `w:abstractNum` and one `w:num`, as fragments.
pub(super) fn numbering_fragments(
    abstract_numbering: &AbstractNumbering,
    instance: &NumberingInstance,
) -> (String, String) {
    let mut writer = XmlWriter::with_capacity(1024);
    write_abstract_numbering(&mut writer, abstract_numbering);
    let abstract_xml = writer.finish();
    let mut writer = XmlWriter::with_capacity(96);
    write_numbering_instance(&mut writer, instance);
    (abstract_xml, writer.finish())
}

/// `xml` with `entries` added: abstract numberings before the root's first
/// `num` or `numIdMacAtCleanup` child, instances before `numIdMacAtCleanup`
/// (CT_Numbering's order), a self-closing root opened for them. An entry whose
/// `numId` the part already has is left out, and the added abstract
/// numberings take the ids after the part's own. The children are found by
/// walking the root's tags, whatever their prefix; the new elements declare
/// `w` themselves when the root uses another prefix. `None` when the part's
/// root is not a well-formed `numbering` element.
pub(super) fn insert_numbering_entries(
    xml: &str,
    entries: &[(AbstractNumbering, NumberingInstance)],
) -> Option<String> {
    let mut at = 0;
    // The prolog: declaration, comments, doctype and whitespace.
    loop {
        at += xml[at..].len() - xml[at..].trim_start().len();
        let rest = &xml[at..];
        at += if rest.starts_with("<?") {
            rest.find("?>")? + 2
        } else if rest.starts_with("<!--") {
            rest.find("-->")? + 3
        } else if rest.starts_with("<!") {
            rest.find('>')? + 1
        } else {
            break;
        };
    }
    let root = tag_name(xml, at)?;
    let (prefix, local) = root.rsplit_once(':').unwrap_or(("", root));
    if local != "numbering" {
        return None;
    }
    let root_end_tag = tag_end(xml, at)?;
    let self_closing = xml[..root_end_tag].ends_with('/');
    let (mut abstract_at, mut num_at) = (None, None);
    let (mut existing, mut last_abstract) = (Vec::new(), None::<f64>);
    let closing = if self_closing {
        root_end_tag
    } else {
        let mut depth = 1;
        let mut cursor = root_end_tag + 1;
        loop {
            let open = cursor + xml[cursor..].find('<')?;
            let rest = &xml[open..];
            if rest.starts_with("<!--") {
                cursor = open + rest.find("-->")? + 3;
            } else if rest.starts_with("<![CDATA[") {
                cursor = open + rest.find("]]>")? + 3;
            } else if rest.starts_with("<?") {
                cursor = open + rest.find("?>")? + 2;
            } else if rest.starts_with("</") {
                depth -= 1;
                if depth == 0 {
                    break open;
                }
                cursor = open + rest.find('>')? + 1;
            } else {
                let name = tag_name(xml, open)?;
                let end = tag_end(xml, open)?;
                if depth == 1 {
                    let child = name.rsplit_once(':').map_or(name, |(_, local)| local);
                    if child == "num" || child == "numIdMacAtCleanup" {
                        abstract_at.get_or_insert(open);
                    }
                    if child == "numIdMacAtCleanup" {
                        num_at.get_or_insert(open);
                    }
                    if child == "num"
                        && let Some(id) = local_attribute(&xml[open..end], "numId")
                    {
                        existing.push(id.to_owned());
                    }
                    if child == "abstractNum"
                        && let Some(id) = local_attribute(&xml[open..end], "abstractNumId")
                            .and_then(|id| id.trim().parse::<f64>().ok())
                    {
                        last_abstract = Some(last_abstract.map_or(id, |last| last.max(id)));
                    }
                }
                if !xml[..end].ends_with('/') {
                    depth += 1;
                }
                cursor = end + 1;
            }
        }
    };
    let declare = if prefix == "w" {
        String::new()
    } else {
        format!(" xmlns:w=\"{W_NS}\"")
    };
    let (mut abstracts, mut nums) = (String::new(), String::new());
    let mut next_abstract = last_abstract.map(|last| last + 1.0);
    for (abstract_numbering, instance) in entries {
        if existing.contains(&int_attr(Some(instance.num_id))) {
            continue;
        }
        let (mut abstract_numbering, mut instance) = (abstract_numbering.clone(), instance.clone());
        if let Some(id) = next_abstract.as_mut() {
            abstract_numbering.abstract_num_id = *id;
            instance.abstract_num_id = *id;
            *id += 1.0;
        }
        let (abstract_xml, num_xml) = numbering_fragments(&abstract_numbering, &instance);
        abstracts.push_str(&abstract_xml.replacen(
            "<w:abstractNum",
            &format!("<w:abstractNum{declare}"),
            1,
        ));
        nums.push_str(&num_xml.replacen("<w:num", &format!("<w:num{declare}"), 1));
    }
    if self_closing {
        let open_tag = xml[..closing - 1].trim_end();
        return Some(format!(
            "{open_tag}>{abstracts}{nums}</{root}>{}",
            &xml[closing + 1..]
        ));
    }
    let abstract_at = abstract_at.unwrap_or(closing);
    let num_at = num_at.unwrap_or(closing);
    Some(
        [
            &xml[..abstract_at],
            &abstracts,
            &xml[abstract_at..num_at],
            &nums,
            &xml[num_at..],
        ]
        .concat(),
    )
}

/// The qualified name of the tag opening at `open`.
fn tag_name(xml: &str, open: usize) -> Option<&str> {
    let name = xml.get(open + 1..)?;
    let length = name.find(|c: char| c.is_whitespace() || c == '/' || c == '>')?;
    (length > 0).then(|| &name[..length])
}

/// The index of the `>` closing the tag that opens at `open`, quotes skipped.
fn tag_end(xml: &str, open: usize) -> Option<usize> {
    let mut quote = None;
    for (index, byte) in xml.bytes().enumerate().skip(open + 1) {
        match (quote, byte) {
            (None, b'"' | b'\'') => quote = Some(byte),
            (Some(open_quote), _) if byte == open_quote => quote = None,
            (None, b'>') => return Some(index),
            _ => {}
        }
    }
    None
}

/// The value of the attribute named `local` (with any prefix) in `tag`.
fn local_attribute<'a>(tag: &'a str, local: &str) -> Option<&'a str> {
    let mut rest = &tag[tag.find(|c: char| c.is_whitespace())?..];
    loop {
        rest = rest.trim_start();
        let equals = rest.find('=')?;
        let name = rest[..equals].trim();
        let value = rest[equals + 1..].trim_start();
        let quote = value.chars().next()?;
        if quote != '"' && quote != '\'' {
            return None;
        }
        let close = value[1..].find(quote)? + 1;
        if name.rsplit_once(':').map_or(name, |(_, local)| local) == local {
            return Some(&value[1..close]);
        }
        rest = &value[close + 1..];
    }
}

fn write_level(writer: &mut XmlWriter, level: &ListLevel) {
    writer
        .start_element("w:lvl")
        .attribute("w:ilvl", &int_attr(Some(level.ilvl)));
    if let Some(value) = level.start {
        writer
            .start_element("w:start")
            .attribute("w:val", &int_attr(Some(value)))
            .end_element();
    }
    writer
        .start_element("w:numFmt")
        .attribute("w:val", &level.num_fmt)
        .end_element();
    if let Some(value) = level.suffix.as_deref().filter(|value| !value.is_empty()) {
        writer
            .start_element("w:suff")
            .attribute("w:val", value)
            .end_element();
    }
    writer
        .start_element("w:lvlText")
        .attribute("w:val", &level.lvl_text)
        .end_element();
    if let Some(value) = level.lvl_jc.as_deref().filter(|value| !value.is_empty()) {
        writer
            .start_element("w:lvlJc")
            .attribute("w:val", value)
            .end_element();
    }
    if let Some(properties) = level
        .p_pr
        .as_ref()
        .filter(|properties| has_indentation(properties))
    {
        writer.start_element("w:pPr");
        write_indentation(writer, properties);
        writer.end_element();
    }
    writer.end_element();
}

fn write_abstract_numbering(writer: &mut XmlWriter, numbering: &AbstractNumbering) {
    writer.start_element("w:abstractNum").attribute(
        "w:abstractNumId",
        &int_attr(Some(numbering.abstract_num_id)),
    );
    if let Some(value) = numbering
        .multi_level_type
        .as_deref()
        .filter(|value| !value.is_empty())
    {
        writer
            .start_element("w:multiLevelType")
            .attribute("w:val", value)
            .end_element();
    }
    let mut levels: Vec<_> = numbering.levels.iter().collect();
    levels.sort_by(|left, right| left.ilvl.total_cmp(&right.ilvl));
    for level in levels {
        write_level(writer, level);
    }
    // Multilevel types use explicit start and end tags.
    if numbering
        .multi_level_type
        .as_deref()
        .unwrap_or_default()
        .is_empty()
        && numbering.levels.is_empty()
    {
        writer.text("");
    }
    writer.end_element();
}

fn write_numbering_instance(writer: &mut XmlWriter, instance: &NumberingInstance) {
    writer
        .start_element("w:num")
        .attribute("w:numId", &int_attr(Some(instance.num_id)))
        .start_element("w:abstractNumId")
        .attribute("w:val", &int_attr(Some(instance.abstract_num_id)))
        .end_element();
    for level_override in instance.level_overrides.as_deref().unwrap_or_default() {
        writer
            .start_element("w:lvlOverride")
            .attribute("w:ilvl", &int_attr(Some(level_override.ilvl)));
        if let Some(value) = level_override.start_override {
            writer
                .start_element("w:startOverride")
                .attribute("w:val", &int_attr(Some(value)))
                .end_element();
        }
        if let Some(level) = level_override.lvl.as_ref() {
            write_level(writer, level);
        }
        if level_override.start_override.is_none() && level_override.lvl.is_none() {
            writer.text("");
        }
        writer.end_element();
    }
    writer.end_element();
}

fn has_indentation(properties: &ParagraphFormatting) -> bool {
    properties.indent_left.is_some()
        || properties.indent_right.is_some()
        || properties
            .indent_first_line
            .is_some_and(|value| properties.hanging_indent == Some(true) || value != 0.0)
}

fn write_indentation(writer: &mut XmlWriter, properties: &ParagraphFormatting) {
    writer.start_element("w:ind");
    if let Some(value) = properties.indent_left {
        writer.attribute("w:left", &int_attr(Some(value)));
    }
    if let Some(value) = properties.indent_right {
        writer.attribute("w:right", &int_attr(Some(value)));
    }
    if let Some(value) = properties.indent_first_line {
        if properties.hanging_indent == Some(true) {
            writer.attribute("w:hanging", &int_attr(Some(value.abs())));
        } else if value != 0.0 {
            writer.attribute("w:firstLine", &int_attr(Some(value)));
        }
    }
    writer.end_element();
}

#[cfg(test)]
mod tests {
    use crate::numbering::{LevelOverride, NumberingMap};
    use crate::xml::{ParseBudget, ParseLimits};

    use super::*;

    fn level(ilvl: f64, num_fmt: &str, level_text: &str) -> ListLevel {
        ListLevel {
            ilvl,
            start: None,
            num_fmt: num_fmt.to_owned(),
            lvl_text: level_text.to_owned(),
            lvl_jc: None,
            suffix: None,
            p_pr: None,
            r_pr: None,
            lvl_restart: None,
            is_lgl: None,
            legacy: None,
        }
    }

    #[test]
    fn numbering_serialization_preserves_order_and_indentation_bytes() {
        let mut first = level(1.0, "lowerLetter", "%2)");
        first.start = Some(2.5);
        first.suffix = Some("space".to_owned());
        first.lvl_jc = Some("right".to_owned());
        first.p_pr = Some(ParagraphFormatting {
            indent_left: Some(720.0),
            indent_right: Some(10.5),
            indent_first_line: Some(-360.0),
            hanging_indent: Some(true),
            ..ParagraphFormatting::default()
        });
        let numbering = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 4.0,
                multi_level_type: Some("multilevel".to_owned()),
                num_style_link: None,
                style_link: None,
                levels: vec![first, level(0.0, "decimal", "%1.")],
                name: None,
            }],
            nums: vec![NumberingInstance {
                num_id: 7.0,
                abstract_num_id: 4.0,
                level_overrides: Some(vec![LevelOverride {
                    ilvl: 1.0,
                    start_override: Some(3.5),
                    lvl: None,
                }]),
            }],
        };
        assert_eq!(
            serialize_numbering_xml(&numbering),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"4\"><w:multiLevelType w:val=\"multilevel\"/><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"3\"/><w:numFmt w:val=\"lowerLetter\"/><w:suff w:val=\"space\"/><w:lvlText w:val=\"%2)\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"720\" w:right=\"11\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=\"7\"><w:abstractNumId w:val=\"4\"/><w:lvlOverride w:ilvl=\"1\"><w:startOverride w:val=\"4\"/></w:lvlOverride></w:num></w:numbering>"
        );
    }

    #[test]
    fn preserves_explicit_empty_container_spelling() {
        assert_eq!(
            serialize_numbering_xml(&NumberingDefinitions::default()),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:numbering>"
        );
        let numbering = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 0.0,
                multi_level_type: None,
                num_style_link: None,
                style_link: None,
                levels: Vec::new(),
                name: None,
            }],
            nums: vec![NumberingInstance {
                num_id: 1.0,
                abstract_num_id: 0.0,
                level_overrides: Some(vec![LevelOverride {
                    ilvl: 2.0,
                    start_override: None,
                    lvl: None,
                }]),
            }],
        };
        let xml = serialize_numbering_xml(&numbering);
        assert!(xml.contains("<w:abstractNum w:abstractNumId=\"0\"></w:abstractNum>"));
        assert!(xml.contains("<w:lvlOverride w:ilvl=\"2\"></w:lvlOverride>"));
    }

    #[test]
    fn escapes_all_numbering_strings() {
        let attack = "\"/><evil attr='&";
        let mut attacked_level = level(0.0, attack, attack);
        attacked_level.suffix = Some(attack.to_owned());
        attacked_level.lvl_jc = Some(attack.to_owned());
        let numbering = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 0.0,
                multi_level_type: Some(attack.to_owned()),
                num_style_link: None,
                style_link: None,
                levels: vec![attacked_level],
                name: None,
            }],
            nums: Vec::new(),
        };
        let xml = serialize_numbering_xml(&numbering);
        assert!(!xml.contains("<evil"));
        assert!(!xml.contains("attr='"));
        assert_eq!(
            xml.matches("&quot;/&gt;&lt;evil attr=&apos;&amp;").count(),
            5
        );
    }

    #[test]
    fn emitted_numbering_parse_backs_to_the_serialized_model() {
        let mut original_level = level(0.0, "decimal", "%1.");
        original_level.start = Some(2.0);
        original_level.p_pr = Some(ParagraphFormatting {
            indent_left: Some(720.0),
            indent_first_line: Some(-360.0),
            hanging_indent: Some(true),
            ..ParagraphFormatting::default()
        });
        let original = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 0.0,
                multi_level_type: Some("singleLevel".to_owned()),
                num_style_link: None,
                style_link: None,
                levels: vec![original_level],
                name: None,
            }],
            nums: vec![NumberingInstance {
                num_id: 1.0,
                abstract_num_id: 0.0,
                level_overrides: None,
            }],
        };
        let xml = serialize_numbering_xml(&original);
        let limits = ParseLimits::default();
        let NumberingMap { definitions } = crate::numbering::parse_numbering(
            Some(xml.as_bytes()),
            "word/numbering.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        assert_eq!(definitions, original);
    }

    /// The bullet button's definition for `num_id` (abstract 9).
    fn bullet(num_id: f64) -> Vec<(AbstractNumbering, NumberingInstance)> {
        vec![crate::numbering::list_button_numbering(true, num_id, 9.0)]
    }

    /// The root's children in order, as `local@numId` or `local@abstractNumId`.
    fn children(xml: &str) -> Vec<String> {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = crate::xml::parse_xml(xml.as_bytes(), "numbering.xml", &mut budget).unwrap();
        document
            .root()
            .unwrap()
            .child_elements()
            .map(|child| {
                let id = child
                    .attributes
                    .iter()
                    .find(|(name, _)| name.ends_with("numId") || name.ends_with("abstractNumId"))
                    .map_or("", |(_, value)| value.as_str());
                format!("{}@{id}", child.local_name())
            })
            .collect()
    }

    #[test]
    fn list_definitions_go_in_ct_numbering_order_whatever_the_parts_shape() {
        let w = W_NS;
        // A picture bullet, a comment holding "<w:num ", a newline after "<w:num"
        // and a cleanup entry: the new abstract goes before the first num, the
        // new num before the cleanup, its abstract id after the part's own.
        let shaped = format!(
            "<?xml version=\"1.0\"?>\n<w:numbering xmlns:w=\"{w}\"><w:numPicBullet w:numPicBulletId=\"0\"/><!-- <w:num w:numId=\"7\"/> --><w:abstractNum w:abstractNumId=\"4\"/><w:num\n w:numId=\"1\"><w:abstractNumId w:val=\"4\"/></w:num><w:numIdMacAtCleanup w:val=\"1\"/></w:numbering>"
        );
        let saved = insert_numbering_entries(&shaped, &bullet(2.0)).unwrap();
        assert_eq!(
            children(&saved),
            [
                "numPicBullet@",
                "abstractNum@4",
                "abstractNum@5",
                "num@1",
                "num@2",
                "numIdMacAtCleanup@"
            ]
        );

        // A self-closing root is opened.
        let empty = format!("<w:numbering xmlns:w=\"{w}\" />");
        let saved = insert_numbering_entries(&empty, &bullet(1.0)).unwrap();
        assert_eq!(children(&saved), ["abstractNum@9", "num@1"]);
        assert!(saved.ends_with("</w:numbering>"));

        // Another prefix: the new elements declare w themselves.
        let prefixed = format!(
            "<ns0:numbering xmlns:ns0=\"{w}\"><ns0:num ns0:numId=\"1\"><ns0:abstractNumId ns0:val=\"0\"/></ns0:num></ns0:numbering>"
        );
        let saved = insert_numbering_entries(&prefixed, &bullet(2.0)).unwrap();
        assert_eq!(children(&saved), ["abstractNum@9", "num@1", "num@2"]);
        assert!(saved.contains(&format!(
            "<w:abstractNum xmlns:w=\"{w}\" w:abstractNumId=\"9\">"
        )));

        // A num the parser could not read (no abstractNumId) keeps its id.
        let unread = format!("<w:numbering xmlns:w=\"{w}\"><w:num w:numId=\"3\"/></w:numbering>");
        assert_eq!(
            insert_numbering_entries(&unread, &bullet(3.0)).unwrap(),
            unread
        );

        assert_eq!(insert_numbering_entries("<w:styles/>", &bullet(1.0)), None);
    }
}
