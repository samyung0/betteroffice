use std::borrow::Cow;

use crate::types::{
    LeadingInlineBox, ParagraphBlock, ParagraphExtent, ReservationRef, ReservedRun, Run,
};

fn size(paragraph: &ParagraphBlock) -> Option<&LeadingInlineBox> {
    paragraph
        .attrs
        .as_ref()?
        .leading_inline_box
        .as_ref()
        .filter(|size| {
            size.width.is_finite()
                && size.width > 0.0
                && size.height.is_finite()
                && size.height > 0.0
                && (size.width as f32).is_finite()
                && (size.height as f32).is_finite()
        })
}

pub(crate) fn measurement_paragraph(paragraph: &ParagraphBlock) -> Cow<'_, ParagraphBlock> {
    let Some(size) = size(paragraph) else {
        return Cow::Borrowed(paragraph);
    };
    let mut measured = paragraph.clone();
    measured.attrs.as_mut().unwrap().leading_inline_box = None;
    measured.runs.insert(
        0,
        Run::Reserved(ReservedRun {
            reservation: ReservationRef {
                owner: "numbering".to_owned(),
                id: "firstPictureBullet".to_owned(),
            },
            width: size.width,
            height: size.height,
            pm_start: None,
            pm_end: None,
        }),
    );
    Cow::Owned(measured)
}

fn source_range(
    head_run: &mut usize,
    head_char: &mut usize,
    tail_run: &mut usize,
    tail_char: &mut usize,
    run_count: usize,
) {
    if *tail_run == 0 {
        *head_run = run_count;
        *tail_run = run_count;
        *head_char = 0;
        *tail_char = 0;
    } else {
        if *head_run == 0 {
            *head_char = 0;
        }
        *head_run = head_run.saturating_sub(1);
        *tail_run -= 1;
    }
}

pub(crate) fn restore_extent(paragraph: &ParagraphBlock, extent: &mut ParagraphExtent) {
    let Some(size) = size(paragraph) else {
        return;
    };
    let is_rtl = paragraph.attrs.as_ref().and_then(|attrs| attrs.bidi) == Some(true);
    for line in &mut extent.lines {
        let has_box = line.head_run == 0;
        if has_box {
            line.width = (line.width - size.width).max(0.0);
            let offset = if is_rtl {
                &mut line.right_offset
            } else {
                &mut line.left_offset
            };
            *offset = Some(offset.unwrap_or(0.0) + size.width);
            let height = line.line_height;
            line.ascent = line.ascent.max(size.height);
            line.line_height = line.line_height.max(line.ascent + line.descent);
            extent.total_height += line.line_height - height;
        }
        source_range(
            &mut line.head_run,
            &mut line.head_char,
            &mut line.tail_run,
            &mut line.tail_char,
            paragraph.runs.len(),
        );
        for segment in line.segments.iter_mut().flatten() {
            if segment.head_run == 0 {
                if !is_rtl {
                    segment.left_offset += size.width;
                }
                segment.available_width = (segment.available_width - size.width).max(0.0);
                segment.width = (segment.width - size.width).max(0.0);
            }
            source_range(
                &mut segment.head_run,
                &mut segment.head_char,
                &mut segment.tail_run,
                &mut segment.tail_char,
                paragraph.runs.len(),
            );
        }
        // Where the box's cluster drops out, the remaining slices keep their
        // visual orders, so a line holding its slices as its clusters stores them.
        if line
            .cluster_advances
            .iter()
            .flatten()
            .any(|cluster| cluster.run_index == 0)
        {
            line.materialize_bidi_slices();
        }
        if let Some(advances) = &mut line.run_advances {
            advances.retain_mut(|advance| match advance.run_index.checked_sub(1) {
                Some(index) => {
                    advance.run_index = index;
                    true
                }
                None => false,
            });
        }
        if let Some(clusters) = &mut line.cluster_advances {
            clusters.retain_mut(|cluster| match cluster.run_index.checked_sub(1) {
                Some(index) => {
                    cluster.run_index = index;
                    if has_box && !is_rtl {
                        cluster.x_offset = (cluster.x_offset - size.width).max(0.0);
                    }
                    true
                }
                None => false,
            });
        }
        if let Some(slices) = &mut line.bidi_slices {
            slices.retain_mut(|slice| match slice.run_index.checked_sub(1) {
                Some(index) => {
                    slice.run_index = index;
                    true
                }
                None => false,
            });
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{
        BlockId, ParagraphAttrs, RunFormatting, TextRun, TypesetClusterAdvance, TypesetRow,
    };

    /// A line holding its slices as its clusters writes, once the box's
    /// cluster drops out, what the same line with stored slices writes; a line
    /// without the box keeps its compact form.
    #[test]
    fn the_picture_bullet_keeps_the_visual_orders_of_a_compact_line() {
        let paragraph = ParagraphBlock {
            sdt_groups: None,
            id: BlockId::Num(0.0),
            para_id: None,
            runs: vec![Run::Text(TextRun {
                fmt: RunFormatting::default(),
                text: "abcd".to_owned(),
                pm_start: None,
                pm_end: None,
                inline_sdt_widget: None,
            })],
            attrs: Some(ParagraphAttrs {
                leading_inline_box: Some(LeadingInlineBox {
                    width: 12.0,
                    height: 10.0,
                }),
                ..Default::default()
            }),
            pm_start: None,
            pm_end: None,
        };
        let cluster = |run_index, start_char, x_offset| TypesetClusterAdvance {
            run_index,
            start_char,
            end_char: start_char + 1,
            advance: 6.0,
            x_offset,
            bidi_level: 0,
            logical_order: start_char,
        };
        let line = |head_run, clusters| TypesetRow {
            head_run,
            tail_run: 1,
            tail_char: 4,
            width: 24.0,
            ascent: 8.0,
            descent: 2.0,
            line_height: 10.0,
            cluster_advances: Some(clusters),
            bidi_from_clusters: true,
            ..Default::default()
        };
        let mut compact = ParagraphExtent {
            lines: vec![
                line(
                    0,
                    vec![cluster(0, 0, 0.0), cluster(1, 0, 12.0), cluster(1, 1, 18.0)],
                ),
                line(1, vec![cluster(1, 2, 0.0), cluster(1, 3, 6.0)]),
            ],
            total_height: 20.0,
        };
        let mut stored = compact.clone();
        for line in &mut stored.lines {
            line.materialize_bidi_slices();
        }
        restore_extent(&paragraph, &mut compact);
        restore_extent(&paragraph, &mut stored);
        assert_eq!(
            serde_json::to_value(&compact).unwrap(),
            serde_json::to_value(&stored).unwrap()
        );
        assert!(compact.lines[1].bidi_from_clusters);
    }
}
