# excel-vba-image-fit

Excel VBA tool for fitting the currently selected picture inside the cell under the picture's upper-left corner.

## What it does

`FitSelectedPictureToTopLeftCell` resizes and positions one selected picture so it fits neatly inside its target cell:

- Uses the selected picture, not the most recently pasted picture.
- Uses the picture's `TopLeftCell` as the target cell.
- If that cell is part of a merged range, uses the full merged range.
- Preserves the picture's aspect ratio.
- Resizes pictures up or down to fit completely inside the target range.
- Leaves a 5% margin on every side, so the image can use up to 90% of the target width and height.
- Centers the picture horizontally and vertically.
- Sets `Placement = xlMoveAndSize`, which is Excel's "Move and size with cells" option.
- Ignores charts, SmartArt, icons, groups, and other non-picture objects.
- Shows friendly messages instead of raising errors for unsupported selections.

## Recommended workflow

1. Copy an image.
2. Select the target cell.
3. Press <kbd>Ctrl</kbd>+<kbd>V</kbd> to paste the image.
4. Run `FitSelectedPictureToTopLeftCell` while the picture is selected, or reselect that picture later and run the macro.

Because the macro uses the picture's `TopLeftCell`, it can still identify the target cell even if you select or edit other cells before running it.

## Installation in PERSONAL.XLSB

1. Open Excel for Microsoft 365 on Windows.
2. Press <kbd>Alt</kbd>+<kbd>F11</kbd> to open the VBA editor.
3. In the Project Explorer, open `VBAProject (PERSONAL.XLSB)`.
4. Insert a standard module, or import `modPictureFit.bas`.
5. Save `PERSONAL.XLSB`.
6. Assign `FitSelectedPictureToTopLeftCell` to a keyboard shortcut through Excel's Macro dialog.

## Files

- `modPictureFit.bas` contains the macro and helper procedures.
