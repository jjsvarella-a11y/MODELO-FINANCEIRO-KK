# Open XML SDK validation

Requires .NET 8 and NuGet restore (DocumentFormat.OpenXml 3.3.0).

```
dotnet run --project src/OpenXmlValidation/kk_validator.csproj -- MODEL.xlsx results.json
```

Checks schema and SDK semantic constraints targeting Microsoft365. Exit code 1 means validation errors. R2 had 20 font-child-order schema errors, all in `/xl/styles.xml`. `repair_font_order.py` moves the existing bold node before size/color/name/family; font properties and indices are retained.

SDK validation does not execute Excel formulas, test the Excel UI or certify that Excel will open without a repair prompt. Native opening, recalculation, saving and reopening remain necessary acceptance checks.
