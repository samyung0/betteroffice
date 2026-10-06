//! DOCX sessions over `docx_edit::EngineSession`.

use crate::error::{Error, Result};
use crate::session::{Applied, Item};
use crate::types::{Command, Determinism, Entry, Target};

pub(crate) struct DocxSession;

fn pending() -> Error {
    Error::engine("DOCX is not ported yet")
}

impl DocxSession {
    pub fn open(_base: &[u8], _state: Option<&[u8]>) -> Result<Self> {
        Err(pending())
    }
    pub fn state(&self) -> Vec<u8> {
        Vec::new()
    }
    pub fn entries(&mut self) -> Result<Vec<Item>> {
        Err(pending())
    }
    pub fn editable(&mut self) -> Result<Vec<Entry>> {
        Err(pending())
    }
    pub fn apply(&mut self, _command: &Command) -> Result<Applied> {
        Err(pending())
    }
    pub fn locate(&mut self, _id: &str) -> Result<Target> {
        Err(pending())
    }
    pub fn export(&mut self, _determinism: Determinism) -> Result<Vec<u8>> {
        Err(pending())
    }
}

pub(crate) fn seed(_base: &[u8]) -> Result<Vec<u8>> {
    Err(pending())
}
