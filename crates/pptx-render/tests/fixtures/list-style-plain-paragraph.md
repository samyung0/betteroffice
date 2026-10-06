`list-style-plain-paragraph.pptx` is `list-style-bullets.pptx` with two part payloads changed, as PowerPoint 2013+ writes a bulleted placeholder holding a plain paragraph. It contains no private corpus material.

- `ppt/slideLayouts/slideLayout1.xml`: the body placeholder's `a:lvl1pPr` is `marL="228600" indent="-228600"` (the Office theme's level-1 hanging indent) instead of 342900/-342900, so inherited indents differ from the editor's default list indent.
- `ppt/slides/slide1.xml`: "First level" is `<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>`, PowerPoint's markup for a paragraph whose bullet was turned off. The other "Inherited bullets" items still inherit their markers and indents from the layout.

The list tests turn the plain paragraph back into a list item next to "Text follows bullet", which inherits its indents. The plain paragraph must then be saved and drawn with that item's laid-out `marL`/`indent`.
