`list-style-plain-paragraph.pptx` is `list-style-bullets.pptx` changed to test a plain paragraph being listed beside items that inherit their indents. It contains no private corpus material.

- Master: `p:bodyStyle/a:lvl1pPr` is `marL="228600" indent="-228600"`, where PowerPoint 2013+ keeps the level-1 hanging indent. The layout's body `a:lvl1pPr` loses its 342900/-342900, so level-0 body items inherit the master's value.
- "Inherited bullets": "First level" is `<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>`, PowerPoint's markup for a paragraph whose bullet was turned off. "Text follows bullet" still inherits its marker and indents.
- "Positioned bullets" (`ph idx="2"`) has "Plain line" with no `a:pPr` and "Bulleted line" with its own `•`. The layout's `idx="2"` placeholder holds two paragraphs, at 114300/-114300 and 457200/-228600. Layout paragraphs are looked up by position, so the two slide paragraphs inherit different indents.
- "List table" is a one-cell table whose `a:lstStyle` sets `a:lvl1pPr` 171450/-171450. It holds a `marL="0" indent="0"` + `a:buNone` paragraph and a `•` item. No placeholder or master style reaches a cell.

The list tests list each plain paragraph with its item. The plain paragraph must then be saved and drawn with the item's laid-out `marL`/`indent`.
