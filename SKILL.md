---
name: visio-svg-stencil-importer
description: Batch-create a reusable Microsoft Visio stencil (.vssx) from a folder of SVG icons. Use when a user wants to import SVG icons into a Visio shape library, build a custom stencil, or automate repeated SVG-to-Visio icon-library imports on Windows with desktop Visio and PowerPoint installed.
---

# Visio SVG Stencil Importer

Use `scripts/Import-SvgIconsToVisioStencil.ps1`. It creates one named master shape per SVG and saves a `.vssx` stencil.

Only process icons that the user has obtained lawfully and provided as local files. Do not scrape, capture network traffic, reverse-engineer material-library APIs, bypass access controls, or download icons from third-party platforms. Ask the user to manually download needed assets from an authorized source and preserve the relevant license and attribution.

## Preconditions

- Run on Windows with desktop Microsoft Visio and PowerPoint installed.
- Confirm the source directory contains SVG files.
- Choose a new output path. Do not overwrite an existing stencil unless the user explicitly requests it.

## Run

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Import-SvgIconsToVisioStencil.ps1 `
  -IconDirectory "C:\path\to\icons" `
  -StencilPath "C:\path\to\output\svg-icons.vssx"
```

Use `-Force` only when replacement of the output stencil is authorized. Use `-KeepIntermediate` only for diagnosing Office conversion failures.

The script converts each SVG to vector EMF with PowerPoint because some Visio automation installations reject direct SVG import. It then imports the EMF into a Visio master, preserving vector scaling.

## Verify

After the script exits successfully, verify that the output exists, is non-empty, and contains one Visio master per source SVG. Open it in Visio through **More Shapes > Open Stencil** and drag a representative icon to a drawing page.

Report that the stencil contains vector EMF representations of the SVG artwork, not editable original SVG XML paths. Preserve source SVG attribution and licensing separately.
