use pptx_edit::{DeckSession, snapshot_package};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

#[test]
fn parsed_view_snapshot_matches_an_unedited_editor_session() {
    let package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    let editor = DeckSession::open(FIXTURE, 71).unwrap();

    assert_eq!(
        snapshot_package(&package).unwrap(),
        editor.snapshot().unwrap()
    );
}
