# Pandoc patch notes

Pandoc version patched: **3.7.0.2**  
File modified: `src/Text/Pandoc/Writers/Docx.hs`

---

## Why patch the source?

Pandoc's Word output had two list formatting behaviours that couldn't be overridden via command-line flags, Lua filters, or a reference document template. The only way to fix them was to modify the writer source directly.

---

## Change 1 — Force decimal numbering at all list levels

### Problem

By default, pandoc cycles list numbering styles at nested levels:

```
1. First level (decimal)
   a. Second level (lowerLetter)
      i. Third level (lowerRoman)
```

The knowledge base style guide requires decimal numbering at every level:

```
1. First level
   1.1 Second level
      1.1.1 Third level
```

### Fix

In `src/Text/Pandoc/Writers/Docx.hs`, the `styleFor` function maps list style variants to Word numFmt strings. Each branch was returning a different style. All branches were changed to return `"decimal"`:

```haskell
-- Before (original pandoc behaviour)
styleFor UpperAlpha _   = "upperLetter"
styleFor LowerAlpha _   = "lowerLetter"
styleFor UpperRoman _   = "upperRoman"
styleFor LowerRoman _   = "lowerRoman"
styleFor Decimal _      = "decimal"
styleFor DefaultStyle _ = "decimal"
styleFor _ _            = "bullet"

-- After (patched)
styleFor UpperAlpha _   = "decimal"
styleFor LowerAlpha _   = "decimal"
styleFor UpperRoman _   = "decimal"
styleFor LowerRoman _   = "decimal"
styleFor Decimal _      = "decimal"
styleFor DefaultStyle _ = "decimal"
styleFor _ _            = "decimal"
```

---

## Change 2 — Fix hanging indent alignment for wide step numbers

### Problem

The `hang` value controls the hanging indent width for list markers — the space reserved for the number before the text begins. With the default value, wide step numbers (e.g. `2.22`) would be shifted further right than narrower ones, causing text to misalign across steps in long procedures:

```
1.  Item text starts here
...
1.9  Item text starts here
1.10   Item text starts here   ← shifted right, breaks column alignment
```

### Fix

The `hang` value (in twentieths of a point, i.e. twips) was reduced from the pandoc default to `420`:

```haskell
-- Before
hang = 720  -- (or similar default)

-- After
hang = 420
```

`420` twips provides enough space for two-digit-plus step numbers while keeping the text column consistent across all list items regardless of step number width.

---

## Building the patched binary

Requires [GHC](https://www.haskell.org/ghc/) and [Cabal](https://www.haskell.org/cabal/).

```bash
# Clone pandoc 3.7.0.2
git clone --branch 3.7.0.2 https://github.com/jgm/pandoc.git
cd pandoc

# Replace the writer with the patched version
cp /path/to/pandoc-patch/Docx.hs src/Text/Pandoc/Writers/Docx.hs

# Build (this takes a while)
cabal install --install-method=copy --installdir=./bin

# Verify
./bin/pandoc --version
```

The resulting binary can be placed anywhere on your `PATH` or referenced directly by the converter app.
