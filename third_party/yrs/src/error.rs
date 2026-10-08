use crate::ID;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("{0}")]
    ReadError(#[from] crate::encoding::read::Error),
    #[error("failed to apply update: {0}")]
    UpdateError(#[from] UpdateError),
    #[error("Cannot execute this operation when document garbage collection is set")]
    Gc,
}

#[derive(Debug, Error)]
pub enum UpdateError {
    #[error("block parent {0} must be deleted or shared ref type. Type: {1}")]
    InvalidParent(ID, u8),
    /// Patched for BetterOffice: the integration work passed the budget that
    /// [crate::TransactionMut::limit_work] set (the steps counted so far).
    #[error("integration work passed its budget ({0} steps)")]
    WorkBudgetExceeded(u64),
}
