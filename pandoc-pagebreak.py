#!/usr/bin/env python
import sys
import panflute as pf

def action(elem, doc):
    if isinstance(elem, pf.RawBlock) and 'PAGEBREAK' in elem.text:
        return pf.RawBlock('<w:br w:type="page"/>', format='openxml')

def main(doc=None):
    return pf.run_filter(action, doc=doc)

if __name__ == "__main__":
    main()
