# Dashboard Excel database (projects 2242 and 2094)

The user (Ahmed Batisha, Fichtner) keeps one Excel workbook per project:

| Project | Workbook | GitHub page | Claude original dashboard (slug) | Claude public page (slug, exact title) |
|---|---|---|---|---|
| 2242 – AP418 – NWC 511250274 | `Dashboard_DB_2242_AP418.xlsx` | `2242/index.html` | `51100d92-8cc3-4094-bc97-db02f705d4e6` | `fcdb9ec5-01db-492b-acd9-961c69b00065` – «تنفيذ خط الصرف الصحي الرئيسي الثاني (منى – العزيزية) من منطقة منى إلى محطة معالجة عرنة» |
| 2094 – P1545 – NWC 511250201 | `Dashboard_DB_2094_P1545.xlsx` | `2094/index.html` | `78bf60e4-bb5c-4221-bed3-606c59f7696b` | `b0dc532a-bf66-4b91-8bf5-24953127f5c1` – «إعادة إعمار واستبدال خطوط النقل الدائرية الرئيسية لمياه الشرب بمكة المكرمة (المرحلة الثالثة)» |

He edits the workbook and sends it with «حدّث الداشبورد». Every update goes to all three places:
the GitHub Pages copy, the Claude public page, and the original Claude dashboard (Dashboard type, database-backed).

## Workbook layout
- First sheet «دليل الاستخدام»: instructions, the data date (cell next to «تاريخ البيانات»), and the list of sheets.
- One sheet per dataset. Row 1 (hidden) holds the field keys the sync reads; row 2 the Arabic labels; data from row 3.
- «بيانات المشروع» is key/value (column A hidden key, column C value).
- Hidden sheet `_schema`: dataset id, sheet name, original format (csv / json / sor), column order, types, derived columns, status→Arabic maps.
- Derived on sync (not in the workbook): `status_name` (where the map is one-to-one), `days_to_close`, `days_open`, `month`, `complete` for SOR registers; `n`, `days_open`, `theme` for critical SORs; `kind_name` for work segments.
- Weekly variance and SPI are Excel formulas for weeks after the last reported one (grey cells).
- Not in the workbook (generated from Primavera / KMZ / OSM): daily_plan, wbs, critical_path, plans, plan_factors, manholes, project_route, basemap, landmarks, permit_paths, page_bg.

## Sync steps (for a Claude session)
1. Save the workbook to e.g. `/tmp/db.xlsx`. If it came from a tool that did not calculate formulas, recalculate it first (xlsx skill `recalc.py`).
2. `python3 tools/dashdb/sync.py /tmp/db.xlsx /tmp/out<proj> <proj>/index.html` (add `--check` first to see what changed).
   It patches the embedded data and the «البيانات بتاريخ» line of the GitHub page, writes the files for the original dashboard
   to `/tmp/out<proj>`, plus `changed.json` and `untranslated.json`.
3. English toggle: add an English entry for each string in `untranslated.json` to `const DICT` inside `<script id="i18n">`
   of the page (numbers are written as `#`).
4. Original Claude dashboard: for each dataset in `changed.json`, `upload_asset` the file, then `db_write` merge `datasets/<id>`
   with the new `source.url` (keep `format`: csv for project, weekly, components, invoices; json for the rest).
   `sor_critical` and `sor_themes` share one file `sor_critical.json` ({rows, themes}).
5. Claude public page: `python3 tools/dashdb/make_claude_copy.py <proj>/index.html /tmp/c<proj>.html`, then `publish_files`
   with `shared_live: public` and the exact title above.
6. Commit and push the GitHub page.

`build_excel.py <proj> <page.html> <out.xlsx> <YYYY-MM-DD>` rebuilds a fresh workbook from the page data
(run the xlsx skill `recalc.py` on it afterwards so formula cells have values).
