//! A projection reads every story's bookmarks with one scan of the bookmarks
//! map (`bookmark_positions`) instead of one scan per story.

use docx_edit::{EditingDoc, SegmentContent, seed_from_docx};
use yrs::{Map, ReadTxn, Transact};

const DOCX: &[u8] = include_bytes!(
    "../../betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx"
);

#[test]
fn segments_with_positions_read_once_match_each_story_read() {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, DOCX).unwrap();
    let stories: Vec<String> = {
        let txn = doc.yrs_doc().transact();
        let stories = txn.get_map("stories").unwrap();
        stories.keys(&txn).map(str::to_owned).collect()
    };
    let positions = doc.bookmark_positions();
    let mut marked = 0;
    for story in &stories {
        let segments = doc.story_segments(story).unwrap();
        marked += segments
            .iter()
            .filter(|segment| {
                matches!(&segment.content, SegmentContent::Pilcrow(properties)
                    if properties.values.contains_key("bookmarks"))
            })
            .count();
        assert_eq!(
            doc.story_segments_with(story, &positions).unwrap(),
            segments,
            "{story}"
        );
    }
    assert!(marked > 0, "the fixture's paragraphs carry bookmarks");
    assert!(stories.len() > 1);
}
