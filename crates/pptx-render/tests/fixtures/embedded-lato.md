# Embedded Lato

`embedded-lato.pptx` is `line-spacing.pptx` with Lato embedded the way Google
Slides exports it: `ppt/fonts/Lato-regular.fntdata` and `Lato-bold.fntdata`
(Embedded OpenType 2.2, MicroType Express compressed, from a Google Slides
export; SIL Open Font License 1.1, see
`crates/ooxml-text/tests/fonts/LICENSE-OFL-Lato.txt`) listed in
`p:embeddedFontLst`, and shape 4's first run set in Lato as "Embedded as text
it wraps later".

At 32 pt in the 350 px box the text breaks "Embedded as text | it wraps later"
in Lato and "Embedded as | text it wraps later" in the Liberation Sans
fallback. The Rust tests build their plain, obfuscated and broken variants from
`line-spacing.pptx` at test time; the TypeScript tests open this file.
