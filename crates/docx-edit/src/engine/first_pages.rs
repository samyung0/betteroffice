//! A read-only view's first pages, laid out before the rest of the body, as
//! Google Docs shows its first pages while it lays out the rest (Q801 B).
//! Upstream's prefix pass (#907, #1109, #1111) by name: the body is measured
//! in growing prefixes until one paginates [`MARGIN_PAGES`] past the pages
//! asked for, and those pages are kept.

use docx_layout::measure_blocks::FloatFlow;
use docx_layout::regions::{
    DocumentRegions, RegionSection, advances_region_section, apply_section_geometry_to_blocks,
};
use docx_layout::types::{
    BlockExtent, Input as LayoutInput, LayoutBlock, LayoutOptions, MeasuredBlock, Run, Size,
};

use super::{
    EngineSession, RegionRequest, layout_error_message, margin_side_shape, prepare_measured_arena,
};

/// What [`EngineSession::layout_document_with_regions_first_pages`] laid out.
#[derive(Debug)]
pub struct FirstPages {
    /// Each page's width and height.
    pub sizes: Vec<(f64, f64)>,
    /// The widest page of the whole body, before the rest is laid out.
    pub document_width: f64,
}

/// Pages a prefix lays out past the ones it keeps. What follows the cut moves
/// only pages from the one it lands on: a keep-with-next chain or widow
/// control pulling lines over, a float on its anchor's page, a note continued
/// from the page before.
const MARGIN_PAGES: usize = 2;

/// Body blocks the first prefix measures; each next prefix adds twice as many.
const FIRST_STEP: usize = 32;

/// Characters of body text under which the body is laid out whole: a short
/// file's first pages come no sooner than its whole layout's (Q822 A). The
/// text is what measuring costs, and it is known before anything is measured.
/// Measured 2026-11-22: a 24-page file of 37,000 characters showed its first
/// pages about 80 ms sooner and every page 180 ms later; one of 50,000
/// characters (31 pages) showed them 185 ms sooner.
const MIN_TEXT: usize = 40_000;

/// The page size pagination takes when a request names none.
const DEFAULT_PAGE_SIZE: Size = Size {
    w: 816.0,
    h: 1056.0,
};

/// A measured prefix of the body that lays out past the pages asked for.
struct Cut {
    measured: Vec<MeasuredBlock>,
    /// The options of the body cut inside `section`, which then lays out as
    /// its final section.
    options: LayoutOptions,
    /// The request's section the cut ends in.
    section: usize,
    document_width: f64,
}

impl EngineSession {
    /// The region pass over only as much of the body as fills its first
    /// `pages` pages, each laid out as the whole body lays it out, for a
    /// read-only view that shows them while
    /// [`Self::layout_document_with_regions`] lays out the whole body. Until
    /// then page-count fields (NUMPAGES) show the file's saved result. Retains
    /// those pages for display only.
    ///
    /// `None`, retaining nothing to display, when the body is too short for
    /// the first pages to come sooner ([`MIN_TEXT`]) or no cut leaves them
    /// so: the body ends before it lays out two pages past them, is in
    /// columns from where it would be cut to its end, or holds a section
    /// break inside a content control (which no request section describes) or
    /// a shape wrapped on the side of an inside or outside margin (settled
    /// over the whole body). The caller then lays out the whole body.
    pub fn layout_document_with_regions_first_pages(
        &self,
        input_json: &str,
        pages: usize,
    ) -> Result<Option<FirstPages>, String> {
        let mut request = RegionRequest::parse(input_json)?;
        let (Some(story), Some(env)) = (request.body_story.as_deref(), &request.render_env) else {
            return Ok(None);
        };
        if pages == 0 || request.regions.sections.is_empty() {
            return Ok(None);
        }
        let cut = self
            .with_lowered_story(story, env, |blocks| {
                measure_first_pages(blocks, &request, pages)
            })
            .map_err(|error| error.to_string())??;
        let Some(cut) = cut else {
            return Ok(None);
        };
        request.input.measured = cut.measured;
        request.input.options = cut.options;
        request.regions.sections.truncate(cut.section + 1);
        self.finish_region_pass(input_json, request, None, true)?;
        let mut pagination = self.pagination.borrow_mut();
        let layout = pagination
            .layout
            .as_mut()
            .expect("layout retained after successful pagination");
        // Headers taller than their margin, or notes, can leave fewer pages
        // than the probe counted.
        if layout.pages.len() < pages + MARGIN_PAGES {
            pagination.layout = None;
            return Ok(None);
        }
        layout.pages.truncate(pages);
        Ok(Some(FirstPages {
            sizes: layout
                .pages
                .iter()
                .map(|page| (page.size.w, page.size.h))
                .collect(),
            document_width: cut.document_width,
        }))
    }
}

/// Measures the body as the whole pass does, a growing prefix at a time,
/// until the prefix paginates [`MARGIN_PAGES`] past `pages` outside a section
/// with columns. The floating zones are the whole body's, so every block
/// measures as in the whole pass.
fn measure_first_pages(
    lowered: &[LayoutBlock],
    request: &RegionRequest,
    pages: usize,
) -> Result<Option<Cut>, String> {
    let regions = &request.regions;
    if text_chars(lowered) < MIN_TEXT
        || regions.sections.iter().all(columned)
        || lowered.iter().any(|block| {
            margin_side_shape(block)
                || matches!(block, LayoutBlock::SectionBreak(section) if section.is_self_describing())
        })
    {
        return Ok(None);
    }
    let measurement = &request.measurement;
    let mut blocks = lowered.to_vec();
    let mut options = request.input.options.clone();
    let (widths, geometry) = prepare_measured_arena(&mut blocks, &mut options, regions);
    let mut flow = FloatFlow::new(&blocks, &widths, measurement, Some(&geometry))?;
    let mut step = FIRST_STEP;
    let mut end = 0;
    loop {
        end = prefix_boundary(&blocks, (end + step).min(blocks.len()));
        step *= 2;
        if end == blocks.len() {
            return Ok(None);
        }
        let section = blocks[..end]
            .iter()
            .filter(|block| advances_region_section(block))
            .count();
        // Sections are counted before measuring, so a body whose rest is all
        // in columns is laid out whole without measuring it twice.
        if regions
            .sections
            .get(section..)
            .is_none_or(|rest| rest.iter().all(columned))
        {
            return Ok(None);
        }
        flow.measure_until(&mut blocks, &widths, measurement, end)?;
        if columned(&regions.sections[section]) {
            continue;
        }
        let through = options_through_section(&request.input.options, regions, section);
        // Pagination rewrites the spacing it reads, so the cut takes the
        // measured blocks as they were.
        let measured = measured_prefix(&blocks[..end], flow.extents());
        let mut probe = LayoutInput {
            measured: measured.clone(),
            options: through,
        };
        let probed =
            docx_layout::place::layout_document(&mut probe).map_err(layout_error_message)?;
        if probed.pages.len() >= pages + MARGIN_PAGES {
            return Ok(Some(Cut {
                measured,
                options: probe.options,
                section,
                document_width: widest_page(&blocks, &options),
            }));
        }
    }
}

/// The characters of text in `blocks`, their tables' cells included.
fn text_chars(blocks: &[LayoutBlock]) -> usize {
    blocks
        .iter()
        .map(|block| match block {
            LayoutBlock::Paragraph(paragraph) => paragraph
                .runs
                .iter()
                .map(|run| match run {
                    Run::Text(run) => run.text.chars().count(),
                    _ => 0,
                })
                .sum(),
            LayoutBlock::Table(table) => table
                .rows
                .iter()
                .flat_map(|row| &row.cells)
                .map(|cell| text_chars(&cell.blocks))
                .sum(),
            _ => 0,
        })
        .sum()
}

/// Columns are balanced over their whole section, so no cut ends inside a
/// section with more than one.
fn columned(section: &RegionSection) -> bool {
    section
        .columns
        .as_ref()
        .is_some_and(|columns| columns.count > 1.0)
}

fn measured_prefix(blocks: &[LayoutBlock], extents: &[BlockExtent]) -> Vec<MeasuredBlock> {
    blocks
        .iter()
        .zip(extents)
        .map(|(block, measure)| MeasuredBlock {
            block: block.clone(),
            measure: measure.clone(),
        })
        .collect()
}

/// A block index a measured prefix may end at. The extent of a bare paragraph
/// mark depends on whether a section break follows it, so no prefix ends just
/// before one, or just before a mark that precedes one. A keep-with-next run
/// is placed by its height through its follower, so no prefix ends inside one.
fn prefix_boundary(blocks: &[LayoutBlock], mut end: usize) -> usize {
    let breaks_at = |index: usize| matches!(blocks.get(index), Some(LayoutBlock::SectionBreak(_)));
    let keeps_with_next = |index: usize| {
        matches!(
            blocks.get(index),
            Some(LayoutBlock::Paragraph(paragraph))
                if paragraph.attrs.as_ref().and_then(|attrs| attrs.keep_next) == Some(true)
        )
    };
    while end < blocks.len()
        && (breaks_at(end) || breaks_at(end + 1) || (end > 0 && keeps_with_next(end - 1)))
    {
        end += 1;
    }
    end
}

/// The options of the body cut inside section `last`, which then lays out as
/// its final section: the request's options with the geometry of the sections
/// through `last`.
fn options_through_section(
    request: &LayoutOptions,
    regions: &DocumentRegions,
    last: usize,
) -> LayoutOptions {
    let through = DocumentRegions {
        sections: regions.sections[..=last].to_vec(),
        even_and_odd_headers: regions.even_and_odd_headers,
        ..DocumentRegions::default()
    };
    let mut options = request.clone();
    apply_section_geometry_to_blocks::<LayoutBlock>(&mut [], &mut options, &through);
    options
}

/// The widest page of the whole body, as pagination sizes its pages: the
/// first section's, each section break's and the final section's.
fn widest_page(blocks: &[LayoutBlock], options: &LayoutOptions) -> f64 {
    let first = options.page_size.as_ref().unwrap_or(&DEFAULT_PAGE_SIZE).w;
    let last = options
        .final_page_size
        .as_ref()
        .map_or(first, |size| size.w);
    blocks
        .iter()
        .filter_map(|block| match block {
            LayoutBlock::SectionBreak(section) => section.page_size.as_ref().map(|size| size.w),
            _ => None,
        })
        .fold(first.max(last), f64::max)
}
