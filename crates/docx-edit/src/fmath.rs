//! Drawing and shape math that gives the same bits on every target (#1970):
//! native builds use `libm`, whose results match wasm's, so a page laid out
//! by the server draws as the browser does.

pub(crate) fn cos(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.cos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::cos(x)
    }
}

pub(crate) fn pow(x: f64, y: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.powf(y)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::pow(x, y)
    }
}

pub(crate) fn sin(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.sin()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::sin(x)
    }
}

#[cfg(all(test, not(target_arch = "wasm32")))]
mod tests {
    use super::*;

    #[test]
    fn math_matches_libm_bits() {
        for x in [-3.0, -0.0, 0.0, 0.37, 1.37, 7.0] {
            assert_eq!(sin(x).to_bits(), libm::sin(x).to_bits());
            assert_eq!(cos(x).to_bits(), libm::cos(x).to_bits());
            assert_eq!(pow(7.0, x).to_bits(), libm::pow(7.0, x).to_bits());
        }
    }
}
