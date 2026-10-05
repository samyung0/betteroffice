use std::sync::Arc;

use ooxml_drawingml::{Theme, resolve_color_value_to_hex_with_theme};
use pptx_parse::{RunProperties, TextBody};
use yrs::branch::{Branch, BranchPtr};
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{
    Any, Assoc, IndexedSequence, Map, MapPrelim, MapRef, Out, ReadTxn, StickyIndex, Text,
    TextPrelim, TextRef, Transact, TransactionMut,
};

use crate::model::{source_paragraph_index, validate_xml_text};
use crate::{
    Bullet, BulletFont, CaretAnchor, DeckSession, EditError, EditResult, KIND, LineSpacing,
    PARA_ID, PILCROW_KIND, PROPERTIES_FROM, ParagraphSnapshot, ParagraphSpacing, STORIES,
    StorySnapshot, TextCaps, TextReceipt, TextRunSnapshot, TextStyle, TextStylePatch,
};

const UNDERLINE_TYPES: [&str; 18] = [
    "none",
    "words",
    "sng",
    "dbl",
    "heavy",
    "dotted",
    "dottedHeavy",
    "dash",
    "dashHeavy",
    "dashLong",
    "dashLongHeavy",
    "dotDash",
    "dotDashHeavy",
    "dotDotDash",
    "dotDotDashHeavy",
    "wavy",
    "wavyHeavy",
    "wavyDbl",
];

const ALIGNMENTS: [&str; 7] = ["l", "ctr", "r", "just", "justLow", "dist", "thaiDist"];

const STRIKES: [&str; 3] = ["noStrike", "sngStrike", "dblStrike"];

/// `ST_TextAutonumberScheme`.
const AUTONUMBER_SCHEMES: [&str; 41] = [
    "alphaLcParenBoth",
    "alphaUcParenBoth",
    "alphaLcParenR",
    "alphaUcParenR",
    "alphaLcPeriod",
    "alphaUcPeriod",
    "arabicParenBoth",
    "arabicParenR",
    "arabicPeriod",
    "arabicPlain",
    "romanLcParenBoth",
    "romanUcParenBoth",
    "romanLcParenR",
    "romanUcParenR",
    "romanLcPeriod",
    "romanUcPeriod",
    "circleNumDbPlain",
    "circleNumWdBlackPlain",
    "circleNumWdWhitePlain",
    "arabicDbPeriod",
    "arabicDbPlain",
    "ea1ChsPeriod",
    "ea1ChsPlain",
    "ea1ChtPeriod",
    "ea1ChtPlain",
    "ea1JpnChsDbPeriod",
    "ea1JpnKorPlain",
    "ea1JpnKorPeriod",
    "arabic1Minus",
    "arabic2Minus",
    "hebrew2Minus",
    "thaiAlphaPeriod",
    "thaiAlphaParenR",
    "thaiAlphaParenBoth",
    "thaiNumPeriod",
    "thaiNumParenR",
    "thaiNumParenBoth",
    "hindiAlphaPeriod",
    "hindiNumPeriod",
    "hindiNumParenR",
    "hindiAlpha1Period",
];

/// `ST_TextIndentLevelType`: `lvl` runs 0-8.
pub(crate) const MAX_LEVEL: u32 = 8;
/// One list level, as PowerPoint's Increase List Level shifts `marL`.
pub(crate) const LEVEL_STEP_EMU: i64 = 457_200;
/// The hanging indent a list edit gives its marker, as Google Slides writes.
pub(crate) const LIST_HANG_EMU: i64 = 342_900;
/// The face a character bullet is drawn in, as PowerPoint writes `a:buFont`.
const BULLET_TYPEFACE: &str = "Arial";
/// `ST_TextMargin` and `ST_TextIndent` bounds.
const MAX_MARGIN_EMU: i64 = 51_206_400;
/// `ST_TextSpacingPercentOrPercentString` (as a share) and `ST_TextSpacingPoint`.
const MAX_SPACING_SHARE: f64 = 132.0;
const MAX_SPACING_POINTS: f64 = 1_584.0;

const MARGIN_LEFT: &str = "marginLeft";
const INDENT: &str = "indent";
const LINE_SPACING: &str = "lineSpacingJson";
const SPACE_BEFORE: &str = "spaceBeforeJson";
const SPACE_AFTER: &str = "spaceAfterJson";
const BULLET_FONT: &str = "bulletFontJson";
const BULLET: &str = "bulletJson";
const LEVEL: &str = "level";

/// The run attributes an edit can set; clearing formatting removes them all.
const RUN_ATTRIBUTES: [&str; 11] = [
    "bold",
    "italic",
    "fontSize",
    "color",
    "fontFamily",
    "underline",
    "spacing",
    "baseline",
    "caps",
    "strike",
    "highlight",
];

/// The values land in schema-typed attributes, so junk must fail the edit
/// rather than the file.
pub(crate) fn validate_text_style(style: &TextStyle) -> EditResult<()> {
    validate_style_values(
        style.font_family.as_deref(),
        style.underline.as_deref(),
        style.color.as_deref(),
        style.font_size_pt,
        style.spacing_pt,
        style.baseline_pct,
    )?;
    validate_run_extras(style.strike.as_deref(), style.highlight.as_deref())
}

fn validate_run_extras(strike: Option<&str>, highlight: Option<&str>) -> EditResult<()> {
    if let Some(strike) = strike
        && !STRIKES.contains(&strike)
    {
        return Err(EditError::InvalidText(format!(
            "unrecognized strikethrough {strike:?}"
        )));
    }
    validate_style_values(None, None, highlight, None, None, None)
}

pub(crate) fn validate_style_values(
    font_family: Option<&str>,
    underline: Option<&str>,
    color: Option<&str>,
    font_size_pt: Option<f64>,
    spacing_pt: Option<f64>,
    baseline_pct: Option<f64>,
) -> EditResult<()> {
    if let Some(font_family) = font_family {
        validate_xml_text(font_family)?;
    }
    if let Some(underline) = underline
        && !UNDERLINE_TYPES.contains(&underline)
    {
        return Err(EditError::InvalidText(format!(
            "unrecognized underline type {underline:?}"
        )));
    }
    if let Some(color) = color {
        let rgb = color.strip_prefix('#').unwrap_or(color);
        if rgb.len() != 6 || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(EditError::InvalidText(format!(
                "color {color:?} must be a six-digit hex value"
            )));
        }
    }
    if let Some(size) = font_size_pt
        && (!size.is_finite() || !(1.0..=4_000.0).contains(&size))
    {
        return Err(EditError::InvalidText(format!(
            "font size {size}pt is outside the 1-4000pt range"
        )));
    }
    if let Some(spacing) = spacing_pt
        && (!spacing.is_finite() || !(-4_000.0..=4_000.0).contains(&spacing))
    {
        return Err(EditError::InvalidText(format!(
            "letter spacing {spacing}pt is outside the -4000-4000pt range"
        )));
    }
    if let Some(baseline) = baseline_pct
        && (!baseline.is_finite()
            || !(f64::from(i32::MIN)..=f64::from(i32::MAX)).contains(&(baseline * 1000.0)))
    {
        return Err(EditError::InvalidText(
            "baseline exceeds the signed percentage range".into(),
        ));
    }
    Ok(())
}

fn validate_alignment(alignment: Option<&str>) -> EditResult<()> {
    if let Some(alignment) = alignment
        && !ALIGNMENTS.contains(&alignment)
    {
        return Err(EditError::InvalidText(format!(
            "unrecognized paragraph alignment {alignment:?}"
        )));
    }
    Ok(())
}

pub(crate) fn seed_story(
    stories: &MapRef,
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    body: &TextBody,
    theme: Option<&Theme>,
) -> EditResult<()> {
    let story = stories.insert(txn, story_id, TextPrelim::new(""));
    if body.paragraphs.is_empty() {
        append_pilcrow(&story, txn, &format!("para:{story_id}:0"), None, 0, None);
        return Ok(());
    }
    for (paragraph_index, paragraph) in body.paragraphs.iter().enumerate() {
        for run in &paragraph.runs {
            if !run.text.is_empty() {
                let style = style_from_run_properties(&run.properties, theme);
                let index = story.len(txn);
                insert_styled_text(&story, txn, index, &run.text, &style);
            }
        }
        let bullet_json = paragraph
            .properties
            .bullet
            .as_ref()
            .map(serde_json::to_string)
            .transpose()
            .map_err(|error| EditError::Json(error.to_string()))?;
        append_pilcrow(
            &story,
            txn,
            &format!("para:{story_id}:{paragraph_index}"),
            paragraph.properties.alignment.as_deref(),
            paragraph.properties.level,
            bullet_json.as_deref(),
        );
    }
    Ok(())
}

pub(crate) fn seed_plain_story(
    stories: &MapRef,
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    paragraph_id: &str,
    text: &str,
    style: &TextStyle,
) -> TextRef {
    let story = stories.insert(txn, story_id, TextPrelim::new(""));
    if !text.is_empty() {
        insert_styled_text(&story, txn, 0, text, style);
    }
    append_pilcrow(&story, txn, paragraph_id, None, 0, None);
    story
}

fn append_pilcrow(
    story: &TextRef,
    txn: &mut TransactionMut<'_>,
    paragraph_id: &str,
    alignment: Option<&str>,
    level: u32,
    bullet_json: Option<&str>,
) {
    let index = story.len(txn);
    let pilcrow =
        story.insert_embed_with_attributes(txn, index, MapPrelim::default(), Attrs::default());
    pilcrow.insert(txn, KIND, PILCROW_KIND);
    pilcrow.insert(txn, PARA_ID, paragraph_id);
    pilcrow.insert(txn, "level", level as f64);
    if let Some(alignment) = alignment {
        pilcrow.insert(txn, "alignment", alignment);
    }
    if let Some(bullet_json) = bullet_json {
        pilcrow.insert(txn, "bulletJson", bullet_json);
    }
}

impl DeckSession {
    pub fn story(&self, story_id: &str) -> EditResult<StorySnapshot> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        snapshot_story(&story, &txn, story_id)
    }

    pub fn anchor_caret(&self, story_id: &str, index: u32) -> EditResult<CaretAnchor> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        let length = final_pilcrow_index(&story, &txn)?;
        if index > length {
            return Err(EditError::OutOfBounds { index, length });
        }
        let position = if index == 0 {
            StickyIndex::from_type(&txn, &story, Assoc::Before)
        } else {
            story
                .sticky_index(&txn, index, Assoc::After)
                .ok_or(EditError::OutOfBounds { index, length })?
        };
        Ok(CaretAnchor {
            story_id: story_id.to_owned(),
            position,
        })
    }

    pub fn resolve_caret_anchor(&self, anchor: &CaretAnchor) -> Option<u32> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, &anchor.story_id).ok()?;
        let offset = anchor.position.get_offset(&txn)?;
        let expected = BranchPtr::from(<TextRef as AsRef<Branch>>::as_ref(&story));
        if offset.branch != expected {
            return None;
        }
        let length = final_pilcrow_index(&story, &txn).ok()?;
        Some(offset.index.min(length))
    }

    pub fn insert_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
        text: &str,
        style: &TextStyle,
    ) -> EditResult<TextReceipt> {
        validate_xml_text(text)?;
        validate_text_style(style)?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let final_pilcrow = final_pilcrow_index(&story, &txn)?;
        if index > final_pilcrow {
            return Err(EditError::OutOfBounds {
                index,
                length: final_pilcrow,
            });
        }
        if !text.is_empty() {
            insert_styled_text(&story, &mut txn, index, text, style);
        }
        let length = text.encode_utf16().count() as u32;
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start: index,
            end: index + length,
            text: text.to_owned(),
        })
    }

    pub fn delete_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
    ) -> EditResult<TextReceipt> {
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_range(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        if end > start {
            story.remove_range(&mut txn, start, end - start);
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Replaces `[start, end)` with `text` in one transaction, so a refused
    /// replacement changes nothing. Unlike [`Self::delete_text`], the range
    /// may cross paragraphs: they join, keeping the first one's properties
    /// and id as PowerPoint does. Each `\n` in `text` then splits the
    /// paragraph the way [`Self::insert_paragraph_break`] does.
    pub fn replace_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        text: &str,
        style: &TextStyle,
    ) -> EditResult<TextReceipt> {
        validate_xml_text(text)?;
        validate_text_style(style)?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let final_pilcrow = final_pilcrow_index(&story, &txn)?;
        if start > end || end > final_pilcrow {
            return Err(EditError::OutOfBounds {
                index: end.max(start),
                length: final_pilcrow,
            });
        }
        remove_joining(&story, &mut txn, start, end);
        let mut index = start;
        for (line_index, line) in text.split('\n').enumerate() {
            if line_index > 0 {
                split_paragraph(&story, &mut txn, story_id, index, self.next_id("para"));
                index += 1;
            }
            if !line.is_empty() {
                insert_styled_text(&story, &mut txn, index, line, style);
                index += line.encode_utf16().count() as u32;
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end: index,
            text: text.to_owned(),
        })
    }

    pub fn format_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        patch: &TextStylePatch,
    ) -> EditResult<TextReceipt> {
        validate_style_values(
            patch.font_family.as_deref(),
            patch.underline.as_deref(),
            patch.color.as_deref(),
            patch.font_size_pt,
            patch.spacing_pt,
            patch.baseline_pct,
        )?;
        validate_run_extras(patch.strike.as_deref(), patch.highlight.as_deref())?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        for (segment_start, segment_end) in paragraph_text_segments(&story, &txn, start, end) {
            story.format(
                &mut txn,
                segment_start,
                segment_end - segment_start,
                attrs_from_patch(patch),
            );
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Sets `a:pPr@algn` on every paragraph the range touches; `None` clears
    /// the value so the placeholder cascade applies again.
    pub fn set_paragraph_alignment(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        alignment: Option<&str>,
    ) -> EditResult<TextReceipt> {
        validate_alignment(alignment)?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        let pilcrows = selected_pilcrows(&story, &txn, start, end);
        for pilcrow in pilcrows {
            match alignment {
                Some(alignment) => {
                    pilcrow.insert(&mut txn, "alignment", alignment);
                }
                None => {
                    pilcrow.remove(&mut txn, "alignment");
                }
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Removes the run attributes `only` names (every one an edit can set
    /// when `None`) from `[start, end)`, so the text inherits them again from
    /// its placeholder or list style.
    pub fn clear_text_formatting(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        only: Option<&[String]>,
    ) -> EditResult<TextReceipt> {
        let keys: Vec<&str> = match only {
            Some(only) => only
                .iter()
                .map(|key| {
                    RUN_ATTRIBUTES
                        .iter()
                        .copied()
                        .find(|known| known == key)
                        .ok_or_else(|| {
                            EditError::InvalidText(format!("unknown run attribute {key:?}"))
                        })
                })
                .collect::<EditResult<_>>()?,
            None => RUN_ATTRIBUTES.to_vec(),
        };
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        let cleared: Attrs = keys
            .into_iter()
            .map(|key| (Arc::from(key), Any::Null))
            .collect();
        for (segment_start, segment_end) in paragraph_text_segments(&story, &txn, start, end) {
            story.format(
                &mut txn,
                segment_start,
                segment_end - segment_start,
                cleared.clone(),
            );
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Makes every paragraph the range touches a list item, `levels[lvl % n]`
    /// giving each level its marker, or (`None`) a plain paragraph. Markup
    /// follows PowerPoint: a list item gets `marL`/`indent` with a hanging
    /// marker and `a:buFont` (Arial for a character, the text's for a number);
    /// a plain one gets `a:buNone` and no hanging indent.
    pub fn set_paragraph_list(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        levels: Option<&[Bullet]>,
    ) -> EditResult<TextReceipt> {
        if let Some(levels) = levels {
            validate_list_levels(levels)?;
        }
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        for pilcrow in selected_pilcrows(&story, &txn, start, end) {
            let level = pilcrow_level(&pilcrow, &txn);
            let step = LEVEL_STEP_EMU * i64::from(level);
            match levels {
                Some(levels) => {
                    let bullet = &levels[level as usize % levels.len()];
                    let font = match bullet {
                        Bullet::Character { .. } => {
                            BulletFont::Typeface(BULLET_TYPEFACE.to_owned())
                        }
                        _ => BulletFont::FollowText,
                    };
                    insert_pilcrow_json(&pilcrow, &mut txn, BULLET, bullet)?;
                    insert_pilcrow_json(&pilcrow, &mut txn, BULLET_FONT, &font)?;
                    pilcrow.insert(&mut txn, MARGIN_LEFT, (step + LIST_HANG_EMU) as f64);
                    pilcrow.insert(&mut txn, INDENT, -LIST_HANG_EMU as f64);
                }
                None => {
                    insert_pilcrow_json(&pilcrow, &mut txn, BULLET, &Bullet::None)?;
                    pilcrow.insert(&mut txn, MARGIN_LEFT, step as f64);
                    pilcrow.insert(&mut txn, INDENT, 0_f64);
                }
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Moves every paragraph the range touches `delta` list levels (`lvl`,
    /// 0-8), shifting an explicit `marL` (the paragraph's own or its file
    /// paragraph's) one step per level as PowerPoint does. A paragraph whose
    /// marker is `levels[old % n]` takes `levels[new % n]`; others keep theirs.
    pub fn change_paragraph_level(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        delta: i32,
        levels: Option<&[Bullet]>,
    ) -> EditResult<TextReceipt> {
        if let Some(levels) = levels {
            validate_list_levels(levels)?;
        }
        let source = crate::deck::source_text_body(&self.package, story_id);
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        for pilcrow in selected_pilcrows(&story, &txn, start, end) {
            let old = pilcrow_level(&pilcrow, &txn);
            let new = (i64::from(old) + i64::from(delta)).clamp(0, i64::from(MAX_LEVEL)) as u32;
            if new == old {
                continue;
            }
            let shift = LEVEL_STEP_EMU * (i64::from(new) - i64::from(old));
            let margin = map_number(&pilcrow, &txn, MARGIN_LEFT)
                .map(|value| value as i64)
                .or_else(|| {
                    let paragraph = source_paragraph_of(&pilcrow, &txn, story_id, source)?;
                    paragraph.properties.margin_left
                });
            pilcrow.insert(&mut txn, LEVEL, f64::from(new));
            if let Some(margin) = margin {
                let shifted = (margin + shift).clamp(0, MAX_MARGIN_EMU);
                pilcrow.insert(&mut txn, MARGIN_LEFT, shifted as f64);
            }
            if let Some(levels) = levels {
                let current = map_string(&pilcrow, &txn, BULLET)
                    .and_then(|json| serde_json::from_str::<Bullet>(&json).ok());
                if current.as_ref() == Some(&levels[old as usize % levels.len()]) {
                    let next = &levels[new as usize % levels.len()];
                    insert_pilcrow_json(&pilcrow, &mut txn, BULLET, next)?;
                }
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Sets the line spacing (`a:lnSpc`) and space before/after (`a:spcBef`,
    /// `a:spcAft`) of every paragraph the range touches; `None` leaves one.
    pub fn set_paragraph_spacing(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        spacing: &ParagraphSpacing,
    ) -> EditResult<TextReceipt> {
        for value in [spacing.line, spacing.before, spacing.after]
            .into_iter()
            .flatten()
        {
            validate_spacing(value)?;
        }
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        for pilcrow in selected_pilcrows(&story, &txn, start, end) {
            for (key, value) in [
                (LINE_SPACING, spacing.line),
                (SPACE_BEFORE, spacing.before),
                (SPACE_AFTER, spacing.after),
            ] {
                if let Some(value) = value {
                    insert_pilcrow_json(&pilcrow, &mut txn, key, &value)?;
                }
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Splits the paragraph at `index`. The first half keeps its id; both
    /// halves keep its properties, as PowerPoint continues a list on Enter.
    pub fn insert_paragraph_break(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
    ) -> EditResult<TextReceipt> {
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let final_pilcrow = final_pilcrow_index(&story, &txn)?;
        if index > final_pilcrow {
            return Err(EditError::OutOfBounds {
                index,
                length: final_pilcrow,
            });
        }
        split_paragraph(&story, &mut txn, story_id, index, self.next_id("para"));
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start: index,
            end: index + 1,
            text: "\n".to_owned(),
        })
    }

    /// Joins the paragraph ending at `index` with the next one, keeping the
    /// first one's id and properties.
    pub fn delete_paragraph_break(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
    ) -> EditResult<TextReceipt> {
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let final_pilcrow = final_pilcrow_index(&story, &txn)?;
        if index >= final_pilcrow {
            return Err(EditError::OutOfBounds {
                index,
                length: final_pilcrow,
            });
        }
        let mut offset = 0;
        let is_pilcrow = story.diff(&txn, YChange::identity).into_iter().any(|diff| {
            let length = out_len(&diff.insert);
            let found = offset == index
                && matches!(
                    diff.insert,
                    Out::YMap(ref map)
                        if map_string(map, &txn, KIND).as_deref() == Some(PILCROW_KIND)
                );
            offset += length;
            found
        });
        if !is_pilcrow {
            return Err(EditError::ParagraphBoundary {
                start: index,
                end: index.saturating_add(1),
            });
        }
        remove_joining(&story, &mut txn, index, index + 1);
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start: index,
            end: index + 1,
            text: "\n".to_owned(),
        })
    }
}

pub(crate) fn validate_story<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
) -> EditResult<()> {
    let mut offset = 0;
    let mut pilcrows = Vec::new();
    for diff in story.diff(txn, YChange::identity) {
        let length = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert {
            if map_string(&map, txn, KIND).as_deref() != Some(PILCROW_KIND) {
                return Err(EditError::InvalidState(format!(
                    "story {story_id} contains a non-pilcrow embed"
                )));
            }
            pilcrows.push(offset);
        }
        offset += length;
    }
    if pilcrows.last().copied() != story.len(txn).checked_sub(1) {
        return Err(EditError::InvalidState(format!(
            "story {story_id} has no final pilcrow"
        )));
    }
    Ok(())
}

/// The snapshot `snapshot_story` reads back from a story `seed_story` wrote,
/// computed without materializing a scratch document.
pub(crate) fn baseline_story(
    story_id: &str,
    body: &TextBody,
    theme: Option<&Theme>,
) -> EditResult<StorySnapshot> {
    let mut paragraphs = Vec::with_capacity(body.paragraphs.len().max(1));
    let mut length = 1_u32;
    if body.paragraphs.is_empty() {
        paragraphs.push(ParagraphSnapshot {
            id: format!("para:{story_id}:0"),
            ..ParagraphSnapshot::default()
        });
    } else {
        length = 0;
        for (paragraph_index, paragraph) in body.paragraphs.iter().enumerate() {
            let mut runs: Vec<TextRunSnapshot> = Vec::new();
            let mut last_attrs: Option<TextStyle> = None;
            for run in &paragraph.runs {
                if run.text.is_empty() {
                    continue;
                }
                let style = style_from_run_properties(&run.properties, theme);
                // yrs' format-gap cleanup deletes every marker between two
                // identically styled runs, so their text items squash into a
                // single run in the snapshot.
                if last_attrs.as_ref() == Some(&style) {
                    if let Some(last) = runs.last_mut() {
                        last.text.push_str(&run.text);
                    }
                } else {
                    runs.push(TextRunSnapshot {
                        text: run.text.clone(),
                        style: baseline_style(style.clone()),
                    });
                }
                last_attrs = Some(style);
                length += run.text.encode_utf16().count() as u32;
            }
            length += 1;
            let bullet_json = paragraph
                .properties
                .bullet
                .as_ref()
                .map(serde_json::to_string)
                .transpose()
                .map_err(|error| EditError::Json(error.to_string()))?;
            paragraphs.push(ParagraphSnapshot {
                id: format!("para:{story_id}:{paragraph_index}"),
                alignment: paragraph.properties.alignment.clone(),
                level: paragraph.properties.level,
                bullet_json,
                runs,
                ..ParagraphSnapshot::default()
            });
        }
    }
    Ok(StorySnapshot {
        id: story_id.to_owned(),
        length,
        paragraphs,
    })
}

/// Attribute reads drop non-finite numbers, so the snapshot applies the same
/// normalization the seeded `Any::Number` attributes go through.
fn baseline_style(style: TextStyle) -> TextStyle {
    TextStyle {
        font_size_pt: style.font_size_pt.filter(|value| value.is_finite()),
        spacing_pt: style.spacing_pt.filter(|value| value.is_finite()),
        baseline_pct: style.baseline_pct.filter(|value| value.is_finite()),
        ..style
    }
}

pub(crate) fn snapshot_story<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
) -> EditResult<StorySnapshot> {
    validate_story(story, txn, story_id)?;
    let mut paragraphs = Vec::new();
    let mut runs = Vec::new();
    for diff in story.diff(txn, YChange::identity) {
        match diff.insert {
            Out::Any(Any::String(text)) => runs.push(TextRunSnapshot {
                text: text.to_string(),
                style: style_from_attrs(diff.attributes.as_deref()),
            }),
            Out::YMap(map) => {
                let id = map_string(&map, txn, PARA_ID).unwrap_or_default();
                // A rebase can give a split paragraph a file paragraph of its own.
                let properties_from = map_string(&map, txn, PROPERTIES_FROM)
                    .filter(|_| source_paragraph_index(&id, story_id).is_none());
                paragraphs.push(ParagraphSnapshot {
                    id,
                    alignment: map_string(&map, txn, "alignment"),
                    level: map_number(&map, txn, LEVEL).unwrap_or_default() as u32,
                    bullet_json: map_string(&map, txn, BULLET),
                    properties_from,
                    margin_left: pilcrow_emu(&map, txn, MARGIN_LEFT, 0)?,
                    indent: pilcrow_emu(&map, txn, INDENT, -MAX_MARGIN_EMU)?,
                    line_spacing: pilcrow_spacing(&map, txn, LINE_SPACING)?,
                    space_before: pilcrow_spacing(&map, txn, SPACE_BEFORE)?,
                    space_after: pilcrow_spacing(&map, txn, SPACE_AFTER)?,
                    bullet_font: pilcrow_bullet_font(&map, txn)?,
                    runs: std::mem::take(&mut runs),
                });
            }
            _ => {}
        }
    }
    Ok(StorySnapshot {
        id: story_id.to_owned(),
        length: story.len(txn),
        paragraphs,
    })
}

/// A list edit's marker per level: characters must be short XML text, numbers
/// an `ST_TextAutonumberScheme` with a start in PowerPoint's range.
fn validate_list_levels(levels: &[Bullet]) -> EditResult<()> {
    if levels.is_empty() || levels.len() > (MAX_LEVEL + 1) as usize {
        return Err(EditError::InvalidText(format!(
            "a list needs 1-{} level markers, got {}",
            MAX_LEVEL + 1,
            levels.len()
        )));
    }
    for level in levels {
        match level {
            Bullet::Character { value } => {
                validate_xml_text(value)?;
                if value.trim().is_empty() || value.encode_utf16().count() > 8 {
                    return Err(EditError::InvalidText(format!(
                        "bullet character {value:?} must be 1-8 visible characters"
                    )));
                }
            }
            Bullet::AutoNumber {
                scheme, start_at, ..
            } => {
                if !AUTONUMBER_SCHEMES.contains(&scheme.as_str()) {
                    return Err(EditError::InvalidText(format!(
                        "unrecognized numbering scheme {scheme:?}"
                    )));
                }
                if !(1..=32_767).contains(start_at) {
                    return Err(EditError::InvalidText(format!(
                        "numbering start {start_at} is outside 1-32767"
                    )));
                }
            }
            Bullet::None => {
                return Err(EditError::InvalidText(
                    "a list level needs a marker".to_owned(),
                ));
            }
        }
    }
    Ok(())
}

fn validate_spacing(spacing: LineSpacing) -> EditResult<()> {
    let valid = match spacing {
        LineSpacing::Percent { value } => {
            value.is_finite() && (0.0..=MAX_SPACING_SHARE).contains(&value)
        }
        LineSpacing::Points { value } => {
            value.is_finite() && (0.0..=MAX_SPACING_POINTS).contains(&value)
        }
    };
    if valid {
        Ok(())
    } else {
        Err(EditError::InvalidText(format!(
            "paragraph spacing {spacing:?} is outside PowerPoint's range"
        )))
    }
}

fn pilcrow_level<T: ReadTxn>(pilcrow: &MapRef, txn: &T) -> u32 {
    (map_number(pilcrow, txn, LEVEL).unwrap_or_default() as u32).min(MAX_LEVEL)
}

fn insert_pilcrow_json<T: serde::Serialize>(
    pilcrow: &MapRef,
    txn: &mut TransactionMut<'_>,
    key: &str,
    value: &T,
) -> EditResult<()> {
    let json = serde_json::to_string(value).map_err(|error| EditError::Json(error.to_string()))?;
    pilcrow.insert(txn, key, json);
    Ok(())
}

/// The file paragraph whose markup a pilcrow's paragraph is written with.
fn source_paragraph_of<'a, T: ReadTxn>(
    pilcrow: &MapRef,
    txn: &T,
    story_id: &str,
    source: Option<&'a TextBody>,
) -> Option<&'a pptx_parse::TextParagraph> {
    let id = map_string(pilcrow, txn, PARA_ID)?;
    let index = source_paragraph_index(&id, story_id).or_else(|| {
        source_paragraph_index(&map_string(pilcrow, txn, PROPERTIES_FROM)?, story_id)
    })?;
    source?.paragraphs.get(index)
}

/// An EMU value a peer wrote; out of the schema's range fails the update.
fn pilcrow_emu<T: ReadTxn>(
    pilcrow: &MapRef,
    txn: &T,
    key: &str,
    minimum: i64,
) -> EditResult<Option<i64>> {
    let Some(value) = pilcrow.get(txn, key) else {
        return Ok(None);
    };
    match value {
        Out::Any(Any::Number(number))
            if number.fract() == 0.0
                && (minimum as f64..=MAX_MARGIN_EMU as f64).contains(&number) =>
        {
            Ok(Some(number as i64))
        }
        _ => Err(EditError::InvalidState(format!(
            "paragraph {key} is not an EMU value in range"
        ))),
    }
}

fn pilcrow_spacing<T: ReadTxn>(
    pilcrow: &MapRef,
    txn: &T,
    key: &str,
) -> EditResult<Option<LineSpacing>> {
    let Some(json) = pilcrow_json_string(pilcrow, txn, key)? else {
        return Ok(None);
    };
    let spacing = serde_json::from_str::<LineSpacing>(&json)
        .map_err(|error| EditError::InvalidState(format!("paragraph {key}: {error}")))?;
    validate_spacing(spacing).map_err(|error| EditError::InvalidState(error.to_string()))?;
    Ok(Some(spacing))
}

fn pilcrow_bullet_font<T: ReadTxn>(pilcrow: &MapRef, txn: &T) -> EditResult<Option<BulletFont>> {
    let Some(json) = pilcrow_json_string(pilcrow, txn, BULLET_FONT)? else {
        return Ok(None);
    };
    let font = serde_json::from_str::<BulletFont>(&json)
        .map_err(|error| EditError::InvalidState(format!("paragraph bullet font: {error}")))?;
    if let BulletFont::Typeface(typeface) = &font
        && (typeface.is_empty() || typeface.len() > 256 || validate_xml_text(typeface).is_err())
    {
        return Err(EditError::InvalidState(
            "paragraph bullet font is not a typeface name".to_owned(),
        ));
    }
    Ok(Some(font))
}

fn pilcrow_json_string<T: ReadTxn>(
    pilcrow: &MapRef,
    txn: &T,
    key: &str,
) -> EditResult<Option<String>> {
    match pilcrow.get(txn, key) {
        None => Ok(None),
        Some(Out::Any(Any::String(json))) => Ok(Some(json.to_string())),
        Some(_) => Err(EditError::InvalidState(format!(
            "paragraph {key} is not a string"
        ))),
    }
}

fn story_ref<T: ReadTxn>(txn: &T, story_id: &str) -> EditResult<TextRef> {
    txn.get_map(STORIES)
        .and_then(|stories| stories.get(txn, story_id))
        .and_then(|value| value.cast::<TextRef>().ok())
        .ok_or_else(|| EditError::StoryNotFound(story_id.to_owned()))
}

fn final_pilcrow_index<T: ReadTxn>(story: &TextRef, txn: &T) -> EditResult<u32> {
    validate_story(story, txn, "requested")?;
    Ok(story.len(txn) - 1)
}

fn check_text_range<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> EditResult<()> {
    check_text_bounds(story, txn, start, end)?;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let item_length = out_len(&diff.insert);
        if matches!(diff.insert, Out::YMap(_)) && start <= offset && offset < end {
            return Err(EditError::ParagraphBoundary { start, end });
        }
        offset += item_length;
    }
    Ok(())
}

/// Removes `[start, end)`. Removed pilcrows join the paragraphs around them,
/// and the joined paragraph keeps the first one's properties and id.
fn remove_joining(story: &TextRef, txn: &mut TransactionMut<'_>, start: u32, end: u32) {
    if start == end {
        return;
    }
    let first = pilcrow_from(story, txn, start)
        .filter(|(offset, _)| *offset < end)
        .map(|(_, pilcrow)| pilcrow_properties(&pilcrow, txn));
    story.remove_range(txn, start, end - start);
    let (Some(properties), Some((_, survivor))) = (first, pilcrow_from(story, txn, start)) else {
        return;
    };
    let stale: Vec<String> = pilcrow_properties(&survivor, txn)
        .into_iter()
        .map(|(key, _)| key)
        .filter(|key| properties.iter().all(|(kept, _)| kept != key))
        .collect();
    for key in stale {
        survivor.remove(txn, &key);
    }
    for (key, value) in properties {
        survivor.insert(txn, key, value);
    }
}

/// Splits the paragraph holding `index` with a pilcrow there. The new pilcrow
/// ends the first half and takes the paragraph's id and properties; the old
/// one ends the second half and becomes `new_id`, still taking its file
/// markup from the file paragraph the paragraph is or continues.
fn split_paragraph(
    story: &TextRef,
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    index: u32,
    new_id: String,
) {
    let Some((_, original)) = pilcrow_from(story, txn, index) else {
        return;
    };
    let id = map_string(&original, txn, PARA_ID).unwrap_or_default();
    let in_file = source_paragraph_index(&id, story_id).is_some();
    let first =
        story.insert_embed_with_attributes(txn, index, MapPrelim::default(), Attrs::default());
    first.insert(txn, KIND, PILCROW_KIND);
    for (key, value) in pilcrow_properties(&original, txn) {
        if !(in_file && key == PROPERTIES_FROM) {
            first.insert(txn, key, value);
        }
    }
    if in_file {
        original.insert(txn, PROPERTIES_FROM, id);
    }
    original.insert(txn, PARA_ID, new_id);
}

/// The first pilcrow at or after `index`, with its offset.
fn pilcrow_from<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> Option<(u32, MapRef)> {
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let length = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert
            && offset >= index
        {
            return Some((offset, map));
        }
        offset += length;
    }
    None
}

/// Every pilcrow entry but the kind: the paragraph id and properties.
fn pilcrow_properties<T: ReadTxn>(pilcrow: &MapRef, txn: &T) -> Vec<(String, Any)> {
    pilcrow
        .iter(txn)
        .filter(|(key, _)| *key != KIND)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_owned(), value)),
            _ => None,
        })
        .collect()
}

fn check_text_bounds<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> EditResult<()> {
    let length = story.len(txn);
    if start > end || end > length {
        return Err(EditError::OutOfBounds {
            index: end.max(start),
            length,
        });
    }
    Ok(())
}

/// The pilcrow of every paragraph the range touches. A collapsed caret picks
/// the paragraph it sits in; a range stopping at a paragraph start does not.
fn selected_pilcrows<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> Vec<MapRef> {
    let start = start.min(story.len(txn).saturating_sub(1));
    let mut pilcrows = Vec::new();
    let mut paragraph_start = 0;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let item_length = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert {
            let touches = start <= offset
                && (end > paragraph_start || (start == end && start >= paragraph_start));
            if touches {
                pilcrows.push(map);
            }
            paragraph_start = offset + item_length;
        }
        offset += item_length;
    }
    pilcrows
}

fn paragraph_text_segments<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    start: u32,
    end: u32,
) -> Vec<(u32, u32)> {
    let mut segments = Vec::new();
    let mut paragraph_start = 0;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let item_length = out_len(&diff.insert);
        if matches!(diff.insert, Out::YMap(_)) {
            let segment_start = start.max(paragraph_start);
            let segment_end = end.min(offset);
            if segment_start < segment_end {
                segments.push((segment_start, segment_end));
            }
            paragraph_start = offset + item_length;
        }
        offset += item_length;
    }
    segments
}

fn text_in_range<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> String {
    let mut output = String::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let length = out_len(&diff.insert);
        if let Out::Any(Any::String(text)) = diff.insert {
            let overlap_start = start.saturating_sub(offset).min(length);
            let overlap_end = end.saturating_sub(offset).min(length);
            if overlap_end > overlap_start {
                let utf16 = text.encode_utf16().collect::<Vec<_>>();
                output.push_str(&String::from_utf16_lossy(
                    &utf16[overlap_start as usize..overlap_end as usize],
                ));
            }
        }
        offset += length;
        if offset >= end {
            break;
        }
    }
    output
}

fn insert_styled_text(
    story: &TextRef,
    txn: &mut TransactionMut<'_>,
    index: u32,
    text: &str,
    style: &TextStyle,
) {
    story.insert(txn, index, text);
    let length = text.encode_utf16().count() as u32;
    for (key, value) in style_values(style) {
        story.format(txn, index, length, Attrs::from([(Arc::from(key), value)]));
    }
}

fn style_values(style: &TextStyle) -> [(&'static str, Any); 11] {
    [
        ("bold", style.bold.map(Any::Bool).unwrap_or(Any::Null)),
        ("italic", style.italic.map(Any::Bool).unwrap_or(Any::Null)),
        (
            "fontSize",
            style.font_size_pt.map(Any::Number).unwrap_or(Any::Null),
        ),
        (
            "color",
            style.color.as_deref().map(Any::from).unwrap_or(Any::Null),
        ),
        (
            "fontFamily",
            style
                .font_family
                .as_deref()
                .map(Any::from)
                .unwrap_or(Any::Null),
        ),
        (
            "underline",
            style
                .underline
                .as_deref()
                .map(Any::from)
                .unwrap_or(Any::Null),
        ),
        (
            "spacing",
            style.spacing_pt.map(Any::Number).unwrap_or(Any::Null),
        ),
        (
            "baseline",
            style.baseline_pct.map(Any::Number).unwrap_or(Any::Null),
        ),
        (
            "caps",
            style
                .caps
                .map(|caps| Any::from(caps.as_attribute()))
                .unwrap_or(Any::Null),
        ),
        (
            "strike",
            style.strike.as_deref().map(Any::from).unwrap_or(Any::Null),
        ),
        (
            "highlight",
            style
                .highlight
                .as_deref()
                .map(Any::from)
                .unwrap_or(Any::Null),
        ),
    ]
}

fn attrs_from_patch(patch: &TextStylePatch) -> Attrs {
    let mut attrs = Attrs::default();
    insert_option(&mut attrs, "bold", patch.bold.map(Any::Bool));
    insert_option(&mut attrs, "italic", patch.italic.map(Any::Bool));
    insert_option(&mut attrs, "fontSize", patch.font_size_pt.map(Any::Number));
    insert_option(&mut attrs, "color", patch.color.as_deref().map(Any::from));
    insert_option(
        &mut attrs,
        "fontFamily",
        patch.font_family.as_deref().map(Any::from),
    );
    insert_option(
        &mut attrs,
        "underline",
        patch.underline.as_deref().map(Any::from),
    );
    insert_option(&mut attrs, "spacing", patch.spacing_pt.map(Any::Number));
    insert_option(&mut attrs, "baseline", patch.baseline_pct.map(Any::Number));
    insert_option(&mut attrs, "strike", patch.strike.as_deref().map(Any::from));
    insert_option(
        &mut attrs,
        "highlight",
        patch.highlight.as_deref().map(Any::from),
    );
    attrs
}

fn insert_option(attrs: &mut Attrs, key: &str, value: Option<Any>) {
    if let Some(value) = value {
        attrs.insert(Arc::from(key), value);
    }
}

fn style_from_run_properties(properties: &RunProperties, theme: Option<&Theme>) -> TextStyle {
    TextStyle {
        bold: properties.bold,
        italic: properties.italic,
        font_size_pt: properties.font_size_pt,
        color: resolve_color_value_to_hex_with_theme(properties.color.as_ref(), theme),
        font_family: properties.font_family.clone(),
        underline: properties.underline.clone(),
        spacing_pt: properties.spacing_pt,
        baseline_pct: properties.baseline_pct,
        caps: properties.caps,
        strike: properties.strike.clone(),
        highlight: resolve_color_value_to_hex_with_theme(properties.highlight.as_ref(), theme),
    }
}

fn style_from_attrs(attrs: Option<&Attrs>) -> TextStyle {
    TextStyle {
        bold: attrs.and_then(|attrs| any_bool(attrs.get("bold"))),
        italic: attrs.and_then(|attrs| any_bool(attrs.get("italic"))),
        font_size_pt: attrs.and_then(|attrs| any_number(attrs.get("fontSize"))),
        color: attrs.and_then(|attrs| any_string(attrs.get("color"))),
        font_family: attrs.and_then(|attrs| any_string(attrs.get("fontFamily"))),
        underline: attrs.and_then(|attrs| any_string(attrs.get("underline"))),
        spacing_pt: attrs.and_then(|attrs| any_number(attrs.get("spacing"))),
        baseline_pct: attrs.and_then(|attrs| any_number(attrs.get("baseline"))),
        caps: attrs
            .and_then(|attrs| any_string(attrs.get("caps")))
            .as_deref()
            .and_then(TextCaps::from_attribute),
        strike: attrs.and_then(|attrs| any_string(attrs.get("strike"))),
        highlight: attrs.and_then(|attrs| any_string(attrs.get("highlight"))),
    }
}

fn any_bool(value: Option<&Any>) -> Option<bool> {
    match value {
        Some(Any::Bool(value)) => Some(*value),
        _ => None,
    }
}

fn any_number(value: Option<&Any>) -> Option<f64> {
    match value {
        Some(Any::Number(value)) if value.is_finite() => Some(*value),
        Some(Any::BigInt(value)) => Some(*value as f64),
        _ => None,
    }
}

fn any_string(value: Option<&Any>) -> Option<String> {
    match value {
        Some(Any::String(value)) => Some(value.to_string()),
        _ => None,
    }
}

fn map_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<String> {
    match map.get(txn, key) {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        _ => None,
    }
}

fn map_number<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<f64> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Number(value))) if value.is_finite() => Some(value),
        Some(Out::Any(Any::BigInt(value))) => Some(value as f64),
        _ => None,
    }
}

fn out_len(value: &Out) -> u32 {
    match value {
        Out::Any(Any::String(value)) => value.encode_utf16().count() as u32,
        _ => 1,
    }
}
